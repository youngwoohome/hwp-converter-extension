import { access, mkdir, rename, rm, stat, writeFile } from 'node:fs/promises';
import { constants } from 'node:fs';
import { basename, dirname, extname, join, resolve } from 'node:path';
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

export async function statSafe(filePath: string) {
  try {
    return await stat(filePath);
  } catch {
    return null;
  }
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

export async function writeFileAtomically(filePath: string, contents: string | Buffer): Promise<void> {
  await ensureParentDirectory(filePath);
  const tempPath = join(dirname(filePath), `.${basename(filePath)}.${nanoid(8)}.tmp`);

  try {
    await writeFile(tempPath, contents);
    await rename(tempPath, filePath);
  } catch (error) {
    await rm(tempPath, { force: true }).catch(() => undefined);
    throw error;
  }
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

export function normalizeHwpxOutputPath(filePath: string): string {
  const resolvedPath = normalizeAbsolutePath(filePath);
  return resolvedPath.toLowerCase().endsWith('.hwpx')
    ? resolvedPath
    : `${resolvedPath.replace(/\.[^/.]+$/i, '')}.hwpx`;
}

export async function getUniqueSiblingPath(sourcePath: string, suffix: string, extensionWithDot: string): Promise<string> {
  const ext = extensionWithDot.startsWith('.') ? extensionWithDot : `.${extensionWithDot}`;
  const sourceExt = extname(sourcePath);
  const sourceDir = dirname(sourcePath);
  const sourceBase = sourceExt ? sourcePath.slice(sourceDir.length + 1, -sourceExt.length) : sourcePath.slice(sourceDir.length + 1);

  let candidate = join(sourceDir, `${sourceBase}${suffix}${ext}`);
  let counter = 1;
  while (await fileExists(candidate)) {
    candidate = join(sourceDir, `${sourceBase}${suffix}-${counter}${ext}`);
    counter++;
  }
  return candidate;
}
