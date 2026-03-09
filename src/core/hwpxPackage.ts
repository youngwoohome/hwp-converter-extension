import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { XMLValidator } from 'fast-xml-parser';
import { writeFileAtomically } from '../utils/files.js';

const REQUIRED_ENTRIES = [
  'mimetype',
  'META-INF/container.xml',
  'Contents/content.hpf',
] as const;

const RECOMMENDED_ENTRIES = [
  'version.xml',
  'META-INF/manifest.xml',
  'Contents/header.xml',
  'settings.xml',
] as const;

const REQUIRED_XML_PATTERNS = [
  /^META-INF\/container\.xml$/i,
  /^Contents\/content\.hpf$/i,
  /^Contents\/section\d+\.xml$/i,
] as const;

const RECOMMENDED_XML_PATTERNS = [
  /^version\.xml$/i,
  /^META-INF\/manifest\.xml$/i,
  /^Contents\/header\.xml$/i,
  /^settings\.xml$/i,
] as const;

export interface HwpxPackageValidationSummary {
  sectionEntries: string[];
  assetEntries: string[];
  missingRecommendedEntries: string[];
}

function collectNames(zip: JSZip): string[] {
  return Object.keys(zip.files).filter((name) => !zip.files[name]?.dir);
}

function assertRequiredEntries(names: string[]): void {
  for (const entry of REQUIRED_ENTRIES) {
    if (!names.includes(entry)) {
      throw new Error(`HWPX package is missing required entry: ${entry}`);
    }
  }

  if (!names.some((name) => /^Contents\/section\d+\.xml$/i.test(name))) {
    throw new Error('HWPX package must contain at least one Contents/sectionN.xml part.');
  }
}

async function assertMimetype(zip: JSZip): Promise<void> {
  const file = zip.file('mimetype');
  if (!file) {
    throw new Error('HWPX package is missing required entry: mimetype');
  }

  const mimetype = (await file.async('string')).trim();
  if (mimetype !== 'application/hwp+zip') {
    throw new Error(`Unexpected HWPX mimetype: ${mimetype || '<empty>'}`);
  }
}

async function assertXmlWellFormed(zip: JSZip, names: string[]): Promise<void> {
  for (const name of names) {
    const file = zip.file(name);
    if (!file) continue;

    const needsValidation = REQUIRED_XML_PATTERNS.some((pattern) => pattern.test(name))
      || RECOMMENDED_XML_PATTERNS.some((pattern) => pattern.test(name));
    if (!needsValidation) {
      continue;
    }

    const xml = await file.async('string');
    const validated = XMLValidator.validate(xml);
    if (validated !== true) {
      throw new Error(`Invalid XML in HWPX package entry: ${name}`);
    }
  }
}

export async function validateHwpxPackage(zip: JSZip): Promise<HwpxPackageValidationSummary> {
  const names = collectNames(zip);
  assertRequiredEntries(names);
  await assertMimetype(zip);
  await assertXmlWellFormed(zip, names);

  const missingRecommendedEntries = RECOMMENDED_ENTRIES.filter((entry) => !names.includes(entry));
  const sectionEntries = names.filter((name) => /^Contents\/section\d+\.xml$/i.test(name)).sort();
  const assetEntries = names.filter((name) => /^BinData\//i.test(name)).sort();

  return {
    sectionEntries,
    assetEntries,
    missingRecommendedEntries,
  };
}

export async function loadValidatedHwpxPackage(filePath: string): Promise<{ zip: JSZip; validation: HwpxPackageValidationSummary }> {
  const buffer = await readFile(filePath);
  const zip = await JSZip.loadAsync(buffer);
  const validation = await validateHwpxPackage(zip);
  return { zip, validation };
}

export async function writeValidatedHwpxPackage(filePath: string, zip: JSZip): Promise<HwpxPackageValidationSummary> {
  const output = await zip.generateAsync({ type: 'nodebuffer' });
  const reloaded = await JSZip.loadAsync(output);
  const validation = await validateHwpxPackage(reloaded);
  await writeFileAtomically(filePath, output);
  return validation;
}
