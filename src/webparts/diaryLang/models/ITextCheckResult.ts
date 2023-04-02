import { IWordToAnalyze } from './';

export interface ITextCheckResult {
    incorrectWords: IWordToAnalyze[];
    suggestedText: string;
}