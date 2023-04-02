export interface IOpenAIResponse {
    incorrect_words: { text: string; suggestions: string[] }[];
    suggested_correction: string;
  }