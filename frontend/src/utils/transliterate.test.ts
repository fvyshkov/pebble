import { describe, it, expect } from 'vitest'
import { transliterate } from './transliterate'

describe('transliterate', () => {
  it('transliterates Russian to Latin', () => {
    expect(transliterate('Продукты')).toBe('produkty')
  })

  it('handles spaces as underscores', () => {
    expect(transliterate('Товарные группы')).toBe('tovarnye_gruppy')
  })

  it('preserves latin chars and digits', () => {
    expect(transliterate('test123')).toBe('test123')
  })

  it('handles mixed Russian and Latin', () => {
    expect(transliterate('Группа A')).toBe('gruppa_a')
  })

  it('replaces unknown chars with underscore', () => {
    expect(transliterate('Цена$')).toBe('tsena')
  })

  it('handles empty string', () => {
    expect(transliterate('')).toBe('')
  })

  it('collapses multiple underscores', () => {
    expect(transliterate('А  Б')).toBe('a_b')
  })
})
