import { execFile } from "node:child_process";

const TIMEOUT_MS = 60_000;

export class PowerPointError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "PowerPointError";
  }
}

export async function runAppleScript(script: string): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile("osascript", ["-e", script], { timeout: TIMEOUT_MS }, (err, stdout, stderr) => {
      if (err) {
        const msg = stderr || err.message;
        if (msg.includes("not running") || msg.includes("-600"))
          return reject(new PowerPointError("Microsoft PowerPoint is not running. Please open PowerPoint and try again."));
        if (msg.includes("not allowed") || msg.includes("not permitted"))
          return reject(new PowerPointError("Permission denied. Grant automation access in System Settings > Privacy & Security > Automation."));
        if (msg.includes("timed out") || (err as NodeJS.ErrnoException).code === "ETIMEDOUT")
          return reject(new PowerPointError("AppleScript execution timed out."));
        return reject(new PowerPointError(msg.trim()));
      }
      resolve(stdout.trimEnd());
    });
  });
}

export async function runJxa(script: string): Promise<string> {
  return new Promise((resolve, reject) => {
    execFile("osascript", ["-l", "JavaScript", "-e", script], { timeout: TIMEOUT_MS }, (err, stdout, stderr) => {
      if (err) {
        const msg = stderr || err.message;
        if (msg.includes("not running") || msg.includes("-600"))
          return reject(new PowerPointError("Microsoft PowerPoint is not running. Please open PowerPoint and try again."));
        if (msg.includes("not allowed") || msg.includes("not permitted"))
          return reject(new PowerPointError("Permission denied. Grant automation access in System Settings > Privacy & Security > Automation."));
        if (msg.includes("timed out") || (err as NodeJS.ErrnoException).code === "ETIMEDOUT")
          return reject(new PowerPointError("AppleScript execution timed out."));
        return reject(new PowerPointError(msg.trim()));
      }
      resolve(stdout.trimEnd());
    });
  });
}

export function escapeForAppleScript(str: string): string {
  return str.replace(/\\/g, "\\\\").replace(/"/g, '\\"');
}
