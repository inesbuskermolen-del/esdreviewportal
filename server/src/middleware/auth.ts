import { Request, Response, NextFunction } from 'express'

export interface GIWRequest extends Request {
  giw?: { email: string; isGIW: boolean }
}

/** No-op — auth removed. Sets req.giw for routes that read it (e.g. lastEditedBy). */
export function requireGIW(
  req: GIWRequest,
  _res: Response,
  next: NextFunction,
): void {
  req.giw = { email: 'admin@giw.com.au', isGIW: true }
  next()
}
