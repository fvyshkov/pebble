const MAP: Record<string, string> = {
  а: 'a', б: 'b', в: 'v', г: 'g', д: 'd', е: 'e', ё: 'yo',
  ж: 'zh', з: 'z', и: 'i', й: 'y', к: 'k', л: 'l', м: 'm',
  н: 'n', о: 'o', п: 'p', р: 'r', с: 's', т: 't', у: 'u',
  ф: 'f', х: 'kh', ц: 'ts', ч: 'ch', ш: 'sh', щ: 'shch',
  ъ: '', ы: 'y', ь: '', э: 'e', ю: 'yu', я: 'ya',
}

export function transliterate(text: string): string {
  const result: string[] = []
  for (const ch of text.toLowerCase()) {
    if (ch in MAP) result.push(MAP[ch])
    else if (/[a-z0-9_\- ]/.test(ch)) result.push(ch)
    else result.push('_')
  }
  return result.join('').trim().replace(/[\s\-]+/g, '_').replace(/_+/g, '_').replace(/^_|_$/g, '')
}
