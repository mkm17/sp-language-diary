import { GptContentService } from '../../services/GptContentService';
import { ITextCheckResult, IWordToAnalyze } from '../../models';
import { ConstantsApi } from '../../constants/ConstantsApi';

declare var global: any;

const mockFetch = jest.fn(() => Promise.resolve({
    json: () => Promise.resolve({})
}));

describe('GptContentService', () => {
    afterEach(() => {
        mockFetch.mockClear();
    });

    it('should make a call to the OpenAI API to check the spelling of text', async () => {
        const service = new GptContentService();
        mockFetch.mockResolvedValueOnce({
            json: jest.fn().mockResolvedValueOnce({
                choices: [
                    {
                        message: {
                            content: JSON.stringify({
                                incorrect_words: [
                                    { text: 'wrd1', suggestions: ['word1', 'word2'] },
                                    { text: 'wrd2', suggestions: ['word3', 'word4'] }
                                ],
                                suggested_correction: 'corrected text'
                            })
                        }
                    }
                ]
            })
        });

        global.fetch = mockFetch;

        const result = await service.checkSpelling('some text', 'English');

        expect(mockFetch).toHaveBeenCalledWith('https://api.openai.com/v1/chat/completions', {
            body: JSON.stringify({
                frequency_penalty: 0,
                max_tokens: 2048,
                model: 'gpt-3.5-turbo',
                presence_penalty: 0,
                temperature: 0,
                top_p: 1,
                messages: [{
                    role: 'assistant',
                    content: `find incorrect English words, find maximum 3 suggestions for them,
         show result in JSON format {incorrect_words:[ x:{ 'text', suggestions: []}], suggested_correction:'text' },
         without any additional data: some text`
                }]
            }),
            headers: { 'Authorization': `Bearer ${ConstantsApi.ChatGPTApiKey}`, 'Content-Type': 'application/json' },
            method: 'POST'
        });

        expect(result.incorrectWords).toHaveLength(2);
        expect(result.incorrectWords[0].title).toBe('wrd1');
        expect(result.incorrectWords[0].suggestions).toEqual(['word1', 'word2']);
        expect(result.incorrectWords[1].title).toBe('wrd2');
        expect(result.incorrectWords[1].suggestions).toEqual(['word3', 'word4']);
        expect(result.suggestedText).toBe('corrected text');
    });
    
    describe('getTranslationsForWords', () => {
        
        it('should return an array of words with translations', async () => {
            mockFetch.mockResolvedValueOnce({
                json: jest.fn().mockResolvedValueOnce({
                    choices: [
                        {
                            message: {
                                content: JSON.stringify([
                                    { word: 'hola', translation: 'hello' },
                                    { word: 'adios', translation: 'goodbye' }
                                ])
                            }
                        }
                    ]
                })
            });
    
            global.fetch = mockFetch;
            const service = new GptContentService();
            const words = [
                { title: 'hola', isChecked: true },
                { title: 'adios', isChecked: true }
            ];
            const result = await service.getTranslationsForWords(words, 'es');

            expect(mockFetch).toHaveBeenCalledTimes(1);
            expect(mockFetch).toHaveBeenCalledWith(expect.any(String), {
                body: expect.any(String),
                headers: expect.objectContaining({
                    'Authorization': expect.any(String),
                    'Content-Type': 'application/json'
                }),
                method: 'POST'
            });

            expect(result).toEqual([
                { title: 'hola', isChecked: true, translation: 'hello' },
                { title: 'adios', isChecked: true, translation: 'goodbye' }
            ]);
        });
    });

});