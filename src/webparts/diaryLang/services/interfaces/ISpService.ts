import { IDropdownOption } from 'office-ui-fabric-react';
import { IWordToAnalyze } from '../../models';
import { ILanguageItem } from '../../models/ILanguageItem';

export interface ISpService {
    getLanguagesOptions(): Promise<IDropdownOption[]>;
    getItem(language: string, word: string): Promise<ILanguageItem>;
    getItems(language: string): Promise<ILanguageItem[]>;
    saveItems(analyzedItems: IWordToAnalyze[], language: string): Promise<void>;
    updateItem(itemId: number, skip: boolean, language: string): Promise<void>;
    setWebUrl(url: string): void;
    setDiaryReportData(language: string, totalWords: number, incorrectWords: number, newIncorrectWords: number): Promise<void>;
}