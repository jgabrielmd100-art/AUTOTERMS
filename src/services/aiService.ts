export const analyzeDocument = async (fileBase64: string, mimeType: string, originalText?: string) => {
  try {
    const response = await fetch('/api/analyze-document', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        fileBase64,
        mimeType,
        originalText,
      }),
    });

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error || 'Falha ao analisar documento');
    }

    const data = await response.json();
    return data.text;
  } catch (err) {
    console.error("Erro AI Proxy:", err);
    throw err;
  }
};

