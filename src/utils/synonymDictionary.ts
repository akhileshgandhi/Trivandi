// HOW TO ADD SYNONYMS:
// Key = the word user might type (lowercase)
// Value = array of equivalent terms to also search
// Multi-word synonyms are fine: 'hr': ['human resources']
// Keep keys lowercase. Values can be any case, they will be quoted automatically.

export const SYNONYMS: Record<string, string[]> = {
  // Document types
  'invoice': ['bill', 'receipt'],
  'presentation': ['deck', 'slides', 'ppt'],
  'report': ['analysis', 'summary', 'overview'],
  'contract': ['agreement', 'deal', 'mou'],
  'proposal': ['pitch', 'offer', 'bid'],
  
  // HR / People
  'hr': ['human resources', 'people ops', 'talent'],
  'employee': ['staff', 'team member', 'colleague'],
  'salary': ['compensation', 'pay', 'ctc', 'package'],
  'leave': ['time off', 'vacation', 'holiday', 'pto'],
  'joining': ['onboarding', 'induction'],
  
  // Finance
  'budget': ['forecast', 'plan', 'spend'],
  'expense': ['cost', 'expenditure', 'spend'],
  'purchase': ['procurement', 'buying', 'po'],
  
  // General
  'car': ['vehicle', 'automobile'],
  'meeting': ['call', 'sync', 'standup', 'discussion'],
  'client': ['customer', 'account', 'partner'],
  'project': ['initiative', 'program', 'engagement'],
};

export function expandQuery(query: string): string {
  if (!query) return query;

  const words = query.split(/\s+/);
  const expandedWords = words.map(word => {
    // Strip leading/trailing punctuation that might be attached in KQL, like '(' or ')'
    const match = word.match(/^([^\w]*)([\w-]+)([^\w]*)$/);
    if (match) {
      const prefix = match[1];
      const coreWord = match[2];
      const suffix = match[3];
      const lowerWord = coreWord.toLowerCase();

      // Do not expand if the word is explicitly quoted or escaped in KQL
      if (prefix.includes('"') || suffix.includes('"') || prefix.includes('\\') || suffix.includes('\\')) {
        return word; 
      }

      if (SYNONYMS[lowerWord]) {
        const terms = [lowerWord, ...SYNONYMS[lowerWord]];
        
        // If any term has spaces, we should quote all terms in the group to be safe
        const shouldQuoteAll = terms.some(t => t.includes(' '));
        const finalTerms = terms.map(t => (shouldQuoteAll ? `"${t}"` : t));
        
        return `${prefix}(${finalTerms.join(' OR ')})${suffix}`;
      }
    }
    return word;
  });

  return expandedWords.join(' ');
}
