import { AuditResult } from '../engine/discrepancy';

/**
 * Generates an offline SHA-256 hash from the verified audit results.
 */
export async function generateDigitalSeal(results: AuditResult[]): Promise<string> {
  if (!results || results.length === 0) {
    throw new Error("Cannot generate seal: No audit results provided.");
  }

  // 1. Create a deterministic, stringified version of the payload
  const payloadString = JSON.stringify(results);

  // 2. Encode string to Uint8Array
  const encoder = new TextEncoder();
  const data = encoder.encode(payloadString);

  // 3. Generate hash using native browser Web Crypto API
  const hashBuffer = await crypto.subtle.digest('SHA-256', data);

  // 4. Convert ArrayBuffer to Hex String
  const hashArray = Array.from(new Uint8Array(hashBuffer));
  const hashHex = hashArray.map(b => b.toString(16).padStart(2, '0')).join('');

  return hashHex;
}
