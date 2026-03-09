import { execFile } from 'node:child_process';
import { existsSync } from 'node:fs';
import { cp, mkdir, rm } from 'node:fs/promises';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { promisify } from 'node:util';

const execFileAsync = promisify(execFile);
const rootDir = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const jarPath = join(rootDir, 'jvm-core', 'target', 'hwp-authoritative-core.jar');

function parseArgs(argv) {
  const args = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith('--')) {
      continue;
    }
    const key = token.slice(2);
    const value = argv[index + 1];
    if (!value || value.startsWith('--')) {
      throw new Error(`Missing value for --${key}`);
    }
    args[key] = value;
    index += 1;
  }
  return args;
}

async function copyIfPresent(sourcePath, destPath) {
  if (!existsSync(sourcePath)) {
    return;
  }
  await cp(sourcePath, destPath, { recursive: true });
}

async function installProductionDependencies(outputDir) {
  const commonOptions = {
    cwd: outputDir,
    env: {
      ...process.env,
      CI: '1',
    },
    maxBuffer: 64 * 1024 * 1024,
  };

  if (process.platform === 'win32') {
    await execFileAsync(
      'cmd.exe',
      ['/d', '/s', '/c', 'pnpm install --prod --frozen-lockfile --config.node-linker=hoisted'],
      commonOptions
    );
    return;
  }

  await execFileAsync(
    'pnpm',
    ['install', '--prod', '--frozen-lockfile', '--config.node-linker=hoisted'],
    commonOptions
  );
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const arch = args.arch?.trim();
  const outputDir = args['output-dir'] ? resolve(rootDir, args['output-dir']) : resolve(rootDir, 'release-bundle');

  if (!arch) {
    throw new Error('Missing required --arch argument');
  }

  if (!existsSync(jarPath)) {
    throw new Error(`Missing JVM bundle jar: ${jarPath}. Run "pnpm build:jvm" first.`);
  }

  await rm(outputDir, { recursive: true, force: true });
  await mkdir(outputDir, { recursive: true });

  await copyIfPresent(join(rootDir, 'README.md'), join(outputDir, 'README.md'));
  await copyIfPresent(join(rootDir, '.env.example'), join(outputDir, '.env.example'));
  await copyIfPresent(join(rootDir, 'package.json'), join(outputDir, 'package.json'));
  await copyIfPresent(join(rootDir, 'pnpm-lock.yaml'), join(outputDir, 'pnpm-lock.yaml'));
  await copyIfPresent(join(rootDir, 'dist'), join(outputDir, 'dist'));

  await mkdir(join(outputDir, 'jvm-core', 'target'), { recursive: true });
  await copyIfPresent(jarPath, join(outputDir, 'jvm-core', 'target', 'hwp-authoritative-core.jar'));

  await installProductionDependencies(outputDir);

  console.log(`Prepared release bundle for ${arch} at ${outputDir}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
