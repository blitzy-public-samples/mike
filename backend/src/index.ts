import "dotenv/config";
import express from "express";
import cors from "cors";
import helmet from "helmet";
import rateLimit from "express-rate-limit";
import { chatRouter } from "./routes/chat";
import { projectsRouter } from "./routes/projects";
import { projectChatRouter } from "./routes/projectChat";
import { documentsRouter } from "./routes/documents";
import { tabularRouter } from "./routes/tabular";
import { workflowsRouter } from "./routes/workflows";
import { userRouter } from "./routes/user";
import { downloadsRouter } from "./routes/downloads";
import { caseLawRouter } from "./routes/caseLaw";
// Document compare: isolated redline engine router (create/status/download).
import { comparisonsRouter } from "./routes/comparisons";

const app = express();
const PORT = process.env.PORT ?? 3001;
const isProduction = process.env.NODE_ENV === "production";

function envInt(name: string, fallback: number): number {
  const raw = process.env[name];
  if (!raw) return fallback;
  const parsed = Number.parseInt(raw, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function minutes(value: number): number {
  return value * 60 * 1000;
}

function hours(value: number): number {
  return minutes(value * 60);
}

function makeLimiter(options: {
  windowMs: number;
  max: number;
  message?: string;
}) {
  return rateLimit({
    windowMs: options.windowMs,
    max: options.max,
    standardHeaders: true,
    legacyHeaders: false,
    skip: (req) => req.method === "OPTIONS",
    message: {
      detail:
        options.message ?? "Too many requests. Please try again later.",
    },
  });
}

const generalLimiter = makeLimiter({
  windowMs: minutes(envInt("RATE_LIMIT_GENERAL_WINDOW_MINUTES", 15)),
  max: envInt("RATE_LIMIT_GENERAL_MAX", 300),
});

const chatLimiter = makeLimiter({
  windowMs: minutes(envInt("RATE_LIMIT_CHAT_WINDOW_MINUTES", 15)),
  max: envInt("RATE_LIMIT_CHAT_MAX", 30),
  message: "Too many chat requests. Please try again later.",
});

const chatCreateLimiter = makeLimiter({
  windowMs: minutes(envInt("RATE_LIMIT_CHAT_CREATE_WINDOW_MINUTES", 15)),
  max: envInt("RATE_LIMIT_CHAT_CREATE_MAX", 60),
});

const uploadLimiter = makeLimiter({
  windowMs: hours(envInt("RATE_LIMIT_UPLOAD_WINDOW_HOURS", 1)),
  max: envInt("RATE_LIMIT_UPLOAD_MAX", 50),
  message: "Too many upload requests. Please try again later.",
});

const exportLimiter = makeLimiter({
  windowMs: hours(envInt("RATE_LIMIT_EXPORT_WINDOW_HOURS", 1)),
  max: envInt("RATE_LIMIT_EXPORT_MAX", 10),
  message: "Too many export requests. Please try again later.",
});

const dataDeleteLimiter = makeLimiter({
  windowMs: hours(envInt("RATE_LIMIT_DATA_DELETE_WINDOW_HOURS", 1)),
  max: envInt("RATE_LIMIT_DATA_DELETE_MAX", 20),
  message: "Too many data deletion requests. Please try again later.",
});

// Document compare: the compare run is EXPENSIVE (parses two .docx, runs the
// deterministic diff, emits a redline, and optionally renders via LibreOffice),
// so throttle it like uploads/exports. The redline download is a cheap read,
// so it gets a more generous throttle. Both use OPTIONAL env knobs with
// sensible defaults — no new required secrets.
const compareLimiter = makeLimiter({
  windowMs: hours(envInt("RATE_LIMIT_COMPARE_WINDOW_HOURS", 1)),
  max: envInt("RATE_LIMIT_COMPARE_MAX", 20),
  message: "Too many comparison requests. Please try again later.",
});

const downloadLimiter = makeLimiter({
  windowMs: minutes(envInt("RATE_LIMIT_DOWNLOAD_WINDOW_MINUTES", 15)),
  max: envInt("RATE_LIMIT_DOWNLOAD_MAX", 100),
  message: "Too many download requests. Please try again later.",
});

function jsonLimitForPath(path: string): string {
  return "50mb";
}

app.disable("x-powered-by");
app.set("trust proxy", envInt("TRUST_PROXY_HOPS", 1));

app.use(
  helmet({
    contentSecurityPolicy: {
      directives: {
        defaultSrc: ["'none'"],
        baseUri: ["'none'"],
        frameAncestors: ["'none'"],
      },
    },
    crossOriginEmbedderPolicy: false,
    hsts: isProduction
      ? {
          maxAge: 15552000,
          includeSubDomains: true,
        }
      : false,
    referrerPolicy: { policy: "no-referrer" },
  }),
);

app.use(
  cors({
    origin: process.env.FRONTEND_URL ?? "http://localhost:3000",
    credentials: true,
  }),
);

app.use(generalLimiter);

app.post("/chat", chatLimiter);
app.post("/projects/:projectId/chat", chatLimiter);
app.post("/tabular-review/:reviewId/chat", chatLimiter);
app.post("/tabular-review/:reviewId/generate", chatLimiter);
app.post("/chat/create", chatCreateLimiter);
app.post("/chat/:chatId/generate-title", chatCreateLimiter);
app.post("/single-documents", uploadLimiter);
app.post("/single-documents/:documentId/versions", uploadLimiter);
app.put(
  "/single-documents/:documentId/versions/:versionId/file",
  uploadLimiter,
);
app.post("/projects/:projectId/documents", uploadLimiter);
app.get("/user/export", exportLimiter);
app.get("/user/chats/export", exportLimiter);
app.get("/user/tabular-reviews/export", exportLimiter);
app.delete("/user/account", dataDeleteLimiter);
app.delete("/user/chats", dataDeleteLimiter);
app.delete("/user/projects", dataDeleteLimiter);
app.delete("/user/tabular-reviews", dataDeleteLimiter);

// Document compare: throttle the expensive compare run and the redline download.
app.post("/projects/:projectId/comparisons", compareLimiter);
app.get("/comparisons/:id/download", downloadLimiter);

app.use((req, res, next) =>
  express.json({ limit: jsonLimitForPath(req.path) })(req, res, next),
);

// Document compare / API hardening: the dynamic express.json() parser above
// rejects a malformed or oversized request body by delegating to Express'
// DEFAULT error handler, which in development renders an HTML page containing
// the body-parser stack trace and absolute filesystem paths — an information
// disclosure defect (QA FINAL_ALT F5). body-parser tags the errors it raises
// with a string `type` (e.g. "entity.parse.failed", "entity.too.large") and a
// numeric client-error `status`; translate those into sanitized JSON so no
// implementation details ever reach the client. This 4-arg handler sits between
// the JSON parser and the router mounts: it catches body-parser failures (which
// are raised BEFORE routing) while any non-body-parser error is forwarded
// untouched via next(err), leaving every existing per-router error path
// (e.g. the workflows router's own handler) unchanged.
app.use(
  (
    err: unknown,
    _req: express.Request,
    res: express.Response,
    next: express.NextFunction,
  ) => {
    if (res.headersSent) return next(err);
    const parseError = err as
      | { type?: unknown; status?: unknown; statusCode?: unknown }
      | null;
    const type =
      typeof parseError?.type === "string" ? parseError.type : undefined;
    const status =
      typeof parseError?.status === "number"
        ? parseError.status
        : typeof parseError?.statusCode === "number"
          ? parseError.statusCode
          : undefined;
    // Malformed JSON (the reported defect): body-parser raises a SyntaxError
    // tagged `entity.parse.failed` with status 400.
    if (type === "entity.parse.failed") {
      return void res.status(400).json({ detail: "Malformed JSON body." });
    }
    // Other body-parser client errors (payload too large, unsupported
    // charset/encoding, aborted request, ...) also carry a 4xx status and a
    // `type`; return sanitized JSON instead of leaking a stack trace.
    if (
      type !== undefined &&
      status !== undefined &&
      status >= 400 &&
      status < 500
    ) {
      return void res
        .status(status)
        .json({ detail: "Request body could not be processed." });
    }
    // Not a body-parser error — preserve existing behavior by deferring.
    return next(err);
  },
);

app.use("/chat", chatRouter);
app.use("/projects", projectsRouter);
app.use("/projects/:projectId/chat", projectChatRouter);
app.use("/single-documents", documentsRouter);
app.use("/tabular-review", tabularRouter);
app.use("/workflows", workflowsRouter);
app.use("/user", userRouter);
app.use("/users", userRouter);
app.use("/download", downloadsRouter);
app.use("/case-law", caseLawRouter);
// Document compare: comparisonsRouter defines ABSOLUTE paths spanning two
// route families — project-scoped `POST /projects/:projectId/comparisons` and
// top-level `GET /comparisons/:id` and `GET /comparisons/:id/download`. A single
// base-path mount (like projectChatRouter at "/projects/:projectId/chat") cannot
// serve both families, so mount at the app root. requireAuth is applied per
// route inside the router (matching downloads.ts), so all paths stay protected.
app.use("/", comparisonsRouter);

app.get("/health", (_req, res) => res.json({ ok: true }));

app.listen(PORT, () => {
  console.log(`Mike backend running on port ${PORT}`);
});
