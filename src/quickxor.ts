// QuickXorHash: Microsoft's hash for OneDrive and SharePoint files
// (driveItem.file.hashes.quickXorHash). Ported line for line from the
// published algorithm (github.com/microsoft/quickxorhash; cross-checked
// against the npm package @cloudraker/quickxorhash 1.1.0 — test/files.test.ts
// pins known vectors and a 5 MiB streaming case from that reference).
const QXH_SHIFT = 11;
const QXH_WIDTH_BITS = 160;

export class QuickXorHash {
  private data = new BigUint64Array(3); // 160 bits fit in 2x64 + 32
  private lengthSoFar = 0n;
  private shiftSoFar = 0;

  update(chunk: Uint8Array): this {
    const addSize = chunk.length;
    if (addSize <= 0) return this;
    let vectorArrayIdx = Math.floor(this.shiftSoFar / 64);
    let vectorOffset = this.shiftSoFar % 64;
    const iterations = Math.min(addSize, QXH_WIDTH_BITS);
    for (let i = 0; i < iterations; ++i) {
      const nextCellIdx = vectorArrayIdx === 2 ? 0 : vectorArrayIdx + 1;
      const bitsInCell = vectorArrayIdx === 2 ? 32 : 64;
      let xoredByte = 0;
      for (let j = i; j < addSize; j += QXH_WIDTH_BITS) {
        xoredByte ^= chunk[j];
      }
      const xb = BigInt(xoredByte);
      this.data[vectorArrayIdx] ^= xb << BigInt(vectorOffset);
      if (vectorOffset > bitsInCell - 8) {
        this.data[nextCellIdx] ^= xb >> BigInt(bitsInCell - vectorOffset);
      }
      vectorOffset += QXH_SHIFT;
      if (vectorOffset >= bitsInCell) {
        vectorArrayIdx = nextCellIdx;
        vectorOffset -= bitsInCell;
      }
    }
    this.shiftSoFar = (this.shiftSoFar + QXH_SHIFT * (addSize % QXH_WIDTH_BITS)) % QXH_WIDTH_BITS;
    this.lengthSoFar += BigInt(addSize);
    return this;
  }

  // base64 of the 20-byte little-endian digest, last 8 bytes XORed with the
  // data length — the exact form Graph reports in file.hashes.quickXorHash.
  digestBase64(): string {
    const out = new Uint8Array(20);
    const view = new DataView(out.buffer);
    view.setBigUint64(0, this.data[0], true);
    view.setBigUint64(8, this.data[1], true);
    view.setUint32(16, Number(this.data[2]), true);
    view.setBigUint64(12, view.getBigUint64(12, true) ^ this.lengthSoFar, true);
    return Buffer.from(out).toString("base64");
  }
}
