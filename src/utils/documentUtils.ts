import { Variable } from '../types';

export const extractVariables = (text: string): string[] => {
  // We look for text inside brackets [TEXT]
  const regex = /\[(.*?)\]/g;
  const matches = new Set<string>();
  let match;
  
  // Clean text from common HTML tags and decode &nbsp; to avoid [NOME&nbsp;DA&nbsp;EMPRESA]
  const cleanText = text.replace(/<[^>]*>?/gm, '').replace(/&nbsp;/gi, ' ');
  
  while ((match = regex.exec(cleanText)) !== null) {
    const cleanedVar = match[1].trim().replace(/\s+/g, ' ');
    if (cleanedVar) {
      matches.add(cleanedVar);
    }
  }
  return Array.from(matches);
};

export const fillTemplate = (template: string, variables: Variable[]): string => {
  // Normalize &nbsp; to regular space in the template text before substitution
  let filled = template.replace(/&nbsp;/gi, ' ');
  variables.forEach((v) => {
    const escapedKey = v.key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const regex = new RegExp(`\\[${escapedKey}\\]`, 'g');
    // For rich text, we just replace the markers
    filled = filled.replace(regex, v.value || `[${v.key}]`);
  });
  return filled;
};
