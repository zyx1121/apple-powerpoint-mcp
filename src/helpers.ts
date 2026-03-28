import { PowerPointError } from "./applescript.js";

export function success(data: unknown) {
  return {
    content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }],
  };
}

export function error(message: string) {
  return {
    content: [{ type: "text" as const, text: JSON.stringify({ error: message }) }],
    isError: true as const,
  };
}

export function withErrorHandling<T extends Record<string, unknown>>(
  fn: (args: T) => Promise<ReturnType<typeof success>>,
) {
  return async (args: T) => {
    try {
      return await fn(args);
    } catch (e) {
      if (e instanceof PowerPointError) return error(e.message);
      if (e instanceof Error) return error(e.message);
      return error(String(e));
    }
  };
}
