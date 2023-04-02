import { ServiceKey } from '@microsoft/sp-core-library';
import { ConstantsApi } from '../constants/ConstantsApi';
import { ITextCheckResult, IWordToAnalyze } from '../models';
import { IContentService } from './interfaces';

export class ContentService implements IContentService {
    public static readonly serviceKey: ServiceKey<IContentService> = ServiceKey.create<IContentService>('ContentService', ContentService);

    /**
     * This method takes in a string of text to check for spelling errors and an optional language code parameter
     * @deprecated
     * @param text 
     * @param languageCode 
     * @returns  Promise that resolves to an object of type ITextCheckResult
     */
    public async checkSpelling(text: string, languageCode: string = 'en-US'): Promise<ITextCheckResult> {
        const url = `https://api.bing.microsoft.com/v7.0/spellcheck?mkt=${languageCode}&mode=spell&setlang=${languageCode}`;
        const headers = {
            'Content-Type': 'application/x-www-form-urlencoded',
            'Ocp-Apim-Subscription-Key': ConstantsApi.BingApiKey
        };
        const body = JSON.stringify({ text });
        const options = {
            method: 'POST',
            headers,
            body
        };
        try {
            const response = await fetch(url, options);
            const data = await response.json();

            if (!data || !data.flaggedTokens) { throw Error('Incorrect response format'); }

            return { incorrectWords: data.flaggedTokens, suggestedText: null };
        } catch (error) {
            return Promise.reject(error);
        }
    }

    /**
     * Not implemented
     * @deprecated
     * @param words 
     * @param language 
     * @returns 
     */
    public getTranslationsForWords(words: IWordToAnalyze[], language: string): Promise<IWordToAnalyze[]> {

        return Promise.resolve(words);
    }
}