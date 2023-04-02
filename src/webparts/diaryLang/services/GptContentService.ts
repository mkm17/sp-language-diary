import { ServiceKey } from '@microsoft/sp-core-library';
import { find } from '@microsoft/sp-lodash-subset';
import { ConstantsApi } from '../constants/ConstantsApi';
import { ITextCheckResult, ITranslation, IWordToAnalyze, IOpenAIResponse } from '../models';
import { IContentService } from './interfaces';

export class GptContentService implements IContentService {
    public static readonly serviceKey: ServiceKey<IContentService> = ServiceKey.create<IContentService>('GptContentService', GptContentService);

    /**
     * This method takes in a text string and a language string as inputs and returns a promise of type ITextCheckResult. 
     * The method sends an HTTP POST request to the OpenAI API to check the spelling of the text. 
     * The response from the API is parsed to extract the incorrect words and their suggested corrections 
     * @param text 
     * @param language 
     * @returns an object of type ITextCheckResult containing an array of IWordToAnalyze objects,
     *  where each object represents an incorrect word and its suggested corrections, and a suggestedText string that contains the corrected text
     */
    public async checkSpelling(text: string, language?: string): Promise<ITextCheckResult> {
        try {
            const apiKey = ConstantsApi.ChatGPTApiKey;
            const apiUrl = 'https://api.openai.com/v1/chat/completions';
            const query = `find incorrect ${language} words, find maximum 3 suggestions for them,
         show result in JSON format {incorrect_words:[ x:{ 'text', suggestions: []}], suggested_correction:'text' },
         without any additional data: ${text}`;

            const response = await fetch(apiUrl, {
                body: JSON.stringify(
                    {
                        frequency_penalty: 0,
                        max_tokens: 2048,
                        model: 'gpt-3.5-turbo',
                        presence_penalty: 0,
                        temperature: 0,
                        top_p: 1,
                        messages: [{
                            role: 'assistant',
                            content: query
                        }
                        ]
                    }),
                headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
                method: 'POST'
            });

            const jsonResponse = await response.json();

            const result: IOpenAIResponse = jsonResponse && jsonResponse.choices && jsonResponse.choices.length > 0
                && jsonResponse.choices[0].message &&
                jsonResponse.choices[0].message.content ? JSON.parse(jsonResponse.choices[0].message.content) : null;

            if (!result || !result.incorrect_words) {
                return {
                    incorrectWords: [],
                    suggestedText: null
                };
            }

            const incorrectWords = result.incorrect_words.map(
                (incorrectWord) => {
                    return { title: incorrectWord.text, suggestions: incorrectWord.suggestions } as IWordToAnalyze;
                });
            return {
                incorrectWords: incorrectWords,
                suggestedText: result.suggested_correction
            };
        } catch (e) {
            console.error('Error in checkSpelling', e);
        }
    }

    /**
     * This method takes in an array of IWordToAnalyze objects called words and a language string as inputs, and returns a promise of type IWordToAnalyze[]. The method sends an HTTP POST request to the OpenAI API to translate the words in the words array from the specified language to English. The response from the API is parsed to extract the translations and update the IWordToAnalyze objects in the words array.
     * @param words 
     * @param language 
     * @returns updated words array
     */
    public async getTranslationsForWords(words: IWordToAnalyze[], language?: string): Promise<IWordToAnalyze[]> {
        try {
            const apiKey = ConstantsApi.ChatGPTApiKey;
            const apiUrl = 'https://api.openai.com/v1/chat/completions';

            const wordsToTranslate = words.filter((word) => word.isChecked).map((word) => word.title).join(',');
            const query = `get translations from ${language} to English for words in json format [{"word":"","translation":""}]: ${wordsToTranslate}`;

            const response = await fetch(apiUrl, {
                body: JSON.stringify(
                    {
                        frequency_penalty: 0,
                        max_tokens: 2048,
                        model: 'gpt-3.5-turbo',
                        presence_penalty: 0,
                        temperature: 0,
                        top_p: 1,
                        messages: [{
                            role: 'assistant',
                            content: query
                        }
                        ]
                    }),
                headers: { 'Authorization': `Bearer ${apiKey}`, 'Content-Type': 'application/json' },
                method: 'POST'
            });

            const jsonResponse = await response.json();

            const result: ITranslation[] = jsonResponse && jsonResponse.choices && jsonResponse.choices.length > 0
                && jsonResponse.choices[0].message &&
                jsonResponse.choices[0].message.content ? JSON.parse(jsonResponse.choices[0].message.content) : null;

            if (!result || result.length === 0) {
                return words;
            }

            const wordsWithTranslation = words.map(
                (word) => {
                    if (!!word.translation) { return word; }
                    const translationItem = find(result, (item) => item.word === word.title);

                    return { ...word, translation: translationItem ? translationItem.translation : null } as IWordToAnalyze;
                });

            return wordsWithTranslation;
        }
        catch (e) {
            console.error('Error in getTranslationsForWords', e);
        }
    }
}
