/**
 * Unit Tests for Redact PDF Processor
 * Requirements: 5.1
 * 
 * Tests redaction functionality including validation and area processing.
 */

import { describe, it, expect } from 'vitest';
import { 
  validateRedactionAreas,
  type RedactionArea,
} from '@/lib/pdf/processors/redact';

describe('Redact Processor', () => {
  describe('validateRedactionAreas', () => {
    it('returns invalid when no areas are provided', () => {
      const result = validateRedactionAreas([], 5);

      expect(result.valid).toBe(false);
      expect(result.errors).toContain('No redaction areas specified');
    });

    it('returns valid for correct areas', () => {
      const areas: RedactionArea[] = [
        { page: 1, x: 100, y: 200, width: 150, height: 50 },
        { page: 2, x: 50, y: 100, width: 200, height: 30 },
      ];

      const result = validateRedactionAreas(areas, 5);

      expect(result.valid).toBe(true);
      expect(result.errors).toHaveLength(0);
    });

    it('detects invalid page numbers - page 0', () => {
      const areas: RedactionArea[] = [
        { page: 0, x: 100, y: 200, width: 150, height: 50 },
      ];

      const result = validateRedactionAreas(areas, 5);

      expect(result.valid).toBe(false);
      expect(result.errors.some(e => e.includes('page number'))).toBe(true);
    });

    it('detects invalid page numbers - page exceeds count', () => {
      const areas: RedactionArea[] = [
        { page: 10, x: 50, y: 100, width: 200, height: 30 },
      ];

      const result = validateRedactionAreas(areas, 5);

      expect(result.valid).toBe(false);
      expect(result.errors.some(e => e.includes('page number'))).toBe(true);
    });

    it('detects negative width', () => {
      const areas: RedactionArea[] = [
        { page: 1, x: 100, y: 200, width: -50, height: 50 },
      ];

      const result = validateRedactionAreas(areas, 5);

      expect(result.valid).toBe(false);
      expect(result.errors.some(e => e.includes('Width'))).toBe(true);
    });

    it('detects negative height', () => {
      const areas: RedactionArea[] = [
        { page: 1, x: 100, y: 200, width: 50, height: -50 },
      ];

      const result = validateRedactionAreas(areas, 5);

      expect(result.valid).toBe(false);
      expect(result.errors.some(e => e.includes('Height'))).toBe(true);
    });

    it('detects negative x coordinate', () => {
      const areas: RedactionArea[] = [
        { page: 1, x: -100, y: 200, width: 50, height: 50 },
      ];

      const result = validateRedactionAreas(areas, 5);

      expect(result.valid).toBe(false);
      expect(result.errors.some(e => e.includes('X coordinate'))).toBe(true);
    });

    it('detects negative y coordinate', () => {
      const areas: RedactionArea[] = [
        { page: 1, x: 100, y: -200, width: 50, height: 50 },
      ];

      const result = validateRedactionAreas(areas, 5);

      expect(result.valid).toBe(false);
      expect(result.errors.some(e => e.includes('Y coordinate'))).toBe(true);
    });

    it('reports multiple errors for multiple invalid areas', () => {
      const areas: RedactionArea[] = [
        { page: 0, x: -100, y: -200, width: -50, height: -50 },
      ];

      const result = validateRedactionAreas(areas, 5);

      expect(result.valid).toBe(false);
      expect(result.errors.length).toBeGreaterThan(1);
    });
  });
});
