/**
 * examples/enterprise で使うサンプルロゴ（PNG）を生成する。
 *
 * 外部の画像を持ち込まず、リポジトリ内で再現できるようにするための小さなスクリプト。
 *   node scripts/make-example-logo.mjs
 */
import { deflateSync } from "node:zlib";
import { mkdirSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";

const WIDTH = 160;
const HEIGHT = 40;
const OUT = "examples/enterprise/templates/acme/assets/logo.png";

const NAVY = [16, 54, 95];
const ORANGE = [200, 98, 47];
const WHITE = [255, 255, 255];

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc ^= byte;
    for (let i = 0; i < 8; i += 1) {
      crc = crc & 1 ? (crc >>> 1) ^ 0xedb88320 : crc >>> 1;
    }
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function chunk(type, data) {
  const length = Buffer.alloc(4);
  length.writeUInt32BE(data.length);
  const typeAndData = Buffer.concat([Buffer.from(type, "latin1"), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(typeAndData));
  return Buffer.concat([length, typeAndData, crc]);
}

/** 「■ ACME」風のシンプルなロゴマークを描く */
function pixel(x, y) {
  if (x < 4 || y < 4 || x >= WIDTH - 4 || y >= HEIGHT - 4) return WHITE;
  if (x < 34) return NAVY;
  if (x < 40) return WHITE;
  // 右側に細いアクセントバーを 3 本置く
  const band = Math.floor((x - 40) / 26);
  if (y > HEIGHT / 2 - 6 && y < HEIGHT / 2 + 6 && (x - 40) % 26 < 18 && band < 4) {
    return band === 0 ? NAVY : ORANGE;
  }
  return WHITE;
}

const raw = Buffer.alloc((WIDTH * 3 + 1) * HEIGHT);
let offset = 0;
for (let y = 0; y < HEIGHT; y += 1) {
  raw[offset++] = 0; // filter: none
  for (let x = 0; x < WIDTH; x += 1) {
    const [r, g, b] = pixel(x, y);
    raw[offset++] = r;
    raw[offset++] = g;
    raw[offset++] = b;
  }
}

const ihdr = Buffer.alloc(13);
ihdr.writeUInt32BE(WIDTH, 0);
ihdr.writeUInt32BE(HEIGHT, 4);
ihdr[8] = 8; // bit depth
ihdr[9] = 2; // color type: truecolor
ihdr[10] = 0;
ihdr[11] = 0;
ihdr[12] = 0;

const png = Buffer.concat([
  Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
  chunk("IHDR", ihdr),
  chunk("IDAT", deflateSync(raw, { level: 9 })),
  chunk("IEND", Buffer.alloc(0)),
]);

mkdirSync(dirname(OUT), { recursive: true });
writeFileSync(OUT, png);
console.log(`generated: ${OUT} (${png.length} bytes)`);
