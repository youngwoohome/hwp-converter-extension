import { execFile } from 'node:child_process';
import { access } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), '..', '..');
const defaultJarPath = resolve(rootDir, 'jvm-core', 'target', 'hwp-authoritative-core.jar');
const pomPath = resolve(rootDir, 'jvm-core', 'pom.xml');

let buildPromise: Promise<string> | null = null;

async function fileExists(path: string): Promise<boolean> {
  try {
    await access(path);
    return true;
  } catch {
    return false;
  }
}

function jvmCoreDisabled(): boolean {
  return process.env.HWP_JVM_CORE_DISABLED === '1';
}

function jarPath(): string {
  return process.env.HWP_JVM_CORE_JAR?.trim() || defaultJarPath;
}

async function ensureJarBuilt(): Promise<string> {
  if (jvmCoreDisabled()) {
    throw new Error('JVM HWP core is disabled by HWP_JVM_CORE_DISABLED=1.');
  }

  const resolvedJarPath = jarPath();
  if (await fileExists(resolvedJarPath)) {
    return resolvedJarPath;
  }

  if (!buildPromise) {
    buildPromise = (async () => {
      await execFileAsync('mvn', ['-q', '-DskipTests', '-f', pomPath, 'package'], {
        cwd: rootDir,
        maxBuffer: 64 * 1024 * 1024,
      });
      if (!await fileExists(resolvedJarPath)) {
        throw new Error(`JVM HWP core build completed without producing ${resolvedJarPath}.`);
      }
      return resolvedJarPath;
    })().finally(() => {
      buildPromise = null;
    });
  }

  return buildPromise;
}

async function runJvmCommand(args: string[]): Promise<string> {
  const jar = await ensureJarBuilt();
  const { stdout } = await execFileAsync('java', ['-jar', jar, ...args], {
    cwd: rootDir,
    maxBuffer: 64 * 1024 * 1024,
  });
  return stdout;
}

export async function checkJvmCoreHealth(): Promise<boolean> {
  try {
    const output = await runJvmCommand(['health']);
    return output.trim() === 'ok';
  } catch {
    return false;
  }
}

export async function extractHwpTextWithJvm(inputPath: string): Promise<string> {
  const output = await runJvmCommand(['extract-hwp-text', '--input', inputPath]);
  return output;
}

export async function convertHwpToHwpxWithJvm(inputPath: string, outputPath: string): Promise<void> {
  await runJvmCommand(['convert-hwp-to-hwpx', '--input', inputPath, '--output', outputPath]);
}

export async function createBlankHwpxWithJvm(outputPath: string): Promise<void> {
  await runJvmCommand(['create-blank-hwpx', '--output', outputPath]);
}
