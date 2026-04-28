// ═══════════════════════════════════════════════════════════════════════════
// VESSEL RECONCILIATION SUITE — Cryptographic Integrity Layer
// Uses the Web Crypto API (SubtleCrypto) for a true SHA-256 digest.
// The seal is deterministic: same audit data always produces the same hash.
// ═══════════════════════════════════════════════════════════════════════════

import type { AuditResult } from '../engine/discrepancy';

// ─── Canonical Serialiser ───────────────────────────────────────────────────
// Produces a stable JSON string regardless of object key insertion order.
// Dates are normalised to ISO-8601 UTC strings.

function canonicalise(results: AuditResult[]): string {
  const payload = results.map(r => ({
    c:  r.componentName,
    s:  r.parentSystem,
    oh: r.overhaulDate.toISOString(),
    lh: r.legacyHours,
    vh: r.verifiedHours,
    d:  r.delta,
    ok: r.isCompliant,
  }));
  return JSON.stringify(payload);
}

// ─── SHA-256 Digest ─────────────────────────────────────────────────────────

export async function generateDigitalSeal(results: AuditResult[]): Promise<string> {
  if (results.length === 0) return '0'.repeat(64);

  const encoder  = new TextEncoder();
  const bytes    = encoder.encode(canonicalise(results));
  const hashBuf  = await crypto.subtle.digest('SHA-256', bytes);
  const hashArr  = Array.from(new Uint8Array(hashBuf));

  return hashArr.map(b => b.toString(16).padStart(2, '0')).join('');
}

// ─── Display Helpers ────────────────────────────────────────────────────────

export function truncateSeal(seal: string, prefixLen = 16, suffixLen = 8): string {
  if (seal.length <= prefixLen + suffixLen) return seal;
  return `${seal.slice(0, prefixLen)}…${seal.slice(-suffixLen)}`;
}
