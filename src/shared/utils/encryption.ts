/**
 * Simple encryption/decryption for IDs in the URL to prevent easy manipulation
 */

/**
 * Encrypts a numeric ID to a Base64 string with a simple salt
 */
export const encryptId = (id: string | number): string => {
  if (!id) return '';
  const strId = String(id);
  // Simple Base64 encoding and remove padding (=)
  return btoa(strId).replace(/=/g, '');
};

/**
 * Decrypts a Base64 string back to a numeric ID
 */
export const decryptId = (encryptedId: string): number | null => {
  if (!encryptedId) return null;
  try {
    // Add padding back if necessary for atob
    let base64 = encryptedId;
    while (base64.length % 4 !== 0) {
      base64 += '=';
    }
    const decrypted = atob(base64);
    const id = parseInt(decrypted, 10);
    return isNaN(id) ? null : id;
  } catch (error) {
    // If it's not base64, maybe it's already a raw ID (for backward compatibility)
    const rawId = parseInt(encryptedId, 10);
    return isNaN(rawId) ? null : rawId;
  }
};
