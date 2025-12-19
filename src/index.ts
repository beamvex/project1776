import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import * as xlsx from 'xlsx';

type Args = {
  dir: string;
  message: string;
  date?: string;
  ods?: string;
  sheet?: string;
};

function parseArgs(argv: string[]): Args {
  const out: Partial<Args> = {};
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--dir') out.dir = argv[++i];
    else if (a === '--message') out.message = argv[++i];
    else if (a === '--date') out.date = argv[++i];
    else if (a === '--ods') out.ods = argv[++i];
    else if (a === '--sheet') out.sheet = argv[++i];
  }

  if (!out.dir) out.dir = '../generated-repo';
  if (!out.message) out.message = 'Initial commit';
  return out as Args;
}

function parseDateOrNow(date?: string): Date {
  if (!date) return new Date();
  const d = new Date(date);
  if (Number.isNaN(d.getTime())) {
    throw new Error(
      `Invalid --date. Use an ISO string like 2025-12-19T09:00:00Z or omit --date for now.`
    );
  }

  return d;
}

function runGit(
  cwd: string,
  args: string[],
  env?: Partial<Record<string, string>>
): void {
  execFileSync('git', args, {
    cwd,
    env: {
      ...process.env,
      ...env,
    },
    stdio: 'inherit',
  });
}

function ensureRepo(repoDir: string): void {
  fs.mkdirSync(repoDir, { recursive: true });
  const gitDir = path.join(repoDir, '.git');
  if (!fs.existsSync(gitDir)) {
    runGit(repoDir, ['init']);
  }
}

function readOdsFile(
  odsPath: string,
  sheetName?: string
): Array<Record<string, unknown>> {
  if (!fs.existsSync(odsPath)) {
    throw new Error(`ODS file not found: ${odsPath}`);
  }

  const workbook = xlsx.readFile(odsPath, { cellDates: true });
  const selectedSheetName = sheetName ?? workbook.SheetNames[0];
  if (!selectedSheetName) {
    throw new Error(`No sheets found in ODS file: ${odsPath}`);
  }

  const sheet = workbook.Sheets[selectedSheetName];
  if (!sheet) {
    throw new Error(
      `Sheet not found: ${selectedSheetName}. Available: ${workbook.SheetNames.join(', ')}`
    );
  }

  return xlsx.utils.sheet_to_json(sheet, { defval: null });
}

function main(sundays: string[]): void {
  const args = parseArgs(process.argv.slice(2));
  const repoDir = path.resolve(process.cwd(), args.dir);

  ensureRepo(repoDir);

  if (args.ods) {
    const odsPath = path.resolve(process.cwd(), args.ods);
    const rows = readOdsFile(odsPath, args.sheet);
    let d = 0;
    for (const row of rows) {
      process.stdout.write(row.day + '');
      for (let i = 1; i <= 53; i++) {
        if (row[i]) {
          let theDate = parseDateOrNow(sundays[i - 1]);
          theDate.setDate(theDate.getDate() + d);
          process.stdout.write('x ');
          createFile(repoDir, theDate, `created file for ${row.day} ${i}`);
        } else {
          process.stdout.write('. ');
        }
      }
      process.stdout.write('\n');
      d++;
    }
  }

  //createFile(repoDir, when, 'created a file');

  console.log(`Repo created at: ${repoDir}`);
}

function createFile(repoDir: string, when: Date, message: string): void {
  const outpath = path.join(
    repoDir,
    'src',
    when.getFullYear().toString(),
    when.getMonth().toString(),
    when.getDate().toString(),
    when.getHours().toString(),
    when.getMinutes().toString(),
    when.getSeconds().toString(),
    when.getMilliseconds().toString()
  );
  fs.mkdirSync(outpath, { recursive: true });
  const file = path.join(outpath, 'index.ts');
  const content = `# ${path.basename(repoDir)}\n\nCreated at ${new Date().toISOString()}\n`;
  fs.writeFileSync(file, content, 'utf8');

  runGit(repoDir, ['add', '-A']);

  const dateRfc2822 = when.toUTCString();
  runGit(repoDir, ['commit', '-m', message], {
    GIT_AUTHOR_DATE: dateRfc2822,
    GIT_COMMITTER_DATE: dateRfc2822,
  });
}

const sundays = [];
let tdate = new Date('1976-01-01T09:00:00Z');

do {
  tdate.setDate(tdate.getDate() - 1);
} while (tdate.getDay() > 0);

for (let i = 0; i < 53; i++) {
  sundays.push(tdate.toISOString());
  tdate.setDate(tdate.getDate() + 7);
}

main(sundays);
