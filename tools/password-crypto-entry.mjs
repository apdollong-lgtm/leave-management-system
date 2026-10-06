import {pbkdf2} from '@noble/hashes/pbkdf2.js';
import {sha256} from '@noble/hashes/sha2.js';
export function derive(passwordBytes, saltBytes, iterations) {
  return pbkdf2(sha256, passwordBytes, saltBytes, {c:iterations,dkLen:32});
}
