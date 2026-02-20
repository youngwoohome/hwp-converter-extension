import { access, mkdir, writeFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import { dirname, extname, join, resolve } from 'node:path';
import { tmpdir } from 'node:os';
import { nanoid } from 'nanoid';

export async function fileExists(filePath: string): Promise<boolean> {
  try {
    await access(filePath, constants.F_OK);
    return true;
  } catch {
    return false;
  }
}

export async function ensureParentDirectory(filePath: string): Promise<void> {
  await mkdir(dirname(filePath), { recursive: true });
}

export function normalizeAbsolutePath(filePath: string): string {
  return resolve(filePath);
}

export function extensionOf(filePath: string): string {
  return extname(filePath).toLowerCase();
}

export async function writeTempFile(buffer: Buffer, extensionWithDot: string): Promise<string> {
  const safeExt = extensionWithDot.startsWith('.') ? extensionWithDot : `.${extensionWithDot}`;
  const path = join(tmpdir(), `hwp-converter-${Date.now()}-${nanoid(8)}${safeExt}`);
  await writeFile(path, buffer);
  return path;
}

export function guessSourceFormat(filetype: string | undefined, sourcePath: string): 'hwp' | 'hwpx' | null {
  const normalized = (filetype || '').trim().toLowerCase();
  if (normalized === 'hwp' || normalized === 'hwpx') {
    return normalized;
  }

  const ext = extensionOf(sourcePath);
  if (ext === '.hwp') return 'hwp';
  if (ext === '.hwpx') return 'hwpx';
  return null;
}

export function normalizeTargetFormat(format: string): string {
  return format.trim().toLowerCase();
}
