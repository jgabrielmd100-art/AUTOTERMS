export interface Variable {
  key: string; // The text inside brackets, e.g., "NOME DA EMPRESA"
  value: string;
}

export interface SavedTemplate {
  id: string;
  name: string;
  content: string;
  lastUsed: number;
}

export type AppStep = 'editor' | 'filler' | 'preview';
