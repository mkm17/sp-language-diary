import { ITextCheckResult, IWordToAnalyze } from '../../models';

export interface IContentService {
    checkSpelling(text: string, language?: string): Promise<ITextCheckResult>;
    getTranslationsForWords(words: IWordToAnalyze[], language?: string): Promise<IWordToAnalyze[]>;
}