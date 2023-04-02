import { ContentService } from '../../services/ContentService';
import { ITextCheckResult, IWordToAnalyze } from '../../models';
import { IContentService } from '../../services/interfaces';

declare var global: any;

const mockFetch = jest.fn(() => Promise.resolve({
    json: () => Promise.resolve({})
}));

describe('ContentService', () => {
    
    afterEach(() => {
        mockFetch.mockClear();
    });

    let contentService: IContentService;

    beforeAll(() => {
        contentService = new ContentService();
    });

    describe('checkSpelling', () => {
        it('should return the correct result when given a valid string', async () => {

            mockFetch.mockResolvedValueOnce({
                json: jest.fn().mockResolvedValueOnce(
                    {
                        flaggedTokens: [

                        ]
                    }
                )

            });
            global.fetch = mockFetch;

            // Arrange
            const text = 'This is a test.';
            const expected: ITextCheckResult = {
                incorrectWords: [],
                suggestedText: null
            };

            // Act
            const result = await contentService.checkSpelling(text);

            // Assert
            expect(result).toEqual(expected);
        });

        it('should throw an error when an invalid API key is used', async () => {
            // Arrange
            const text = 'This is a test.';
            const invalidApiKey = 'invalid_api_key';

            mockFetch.mockResolvedValueOnce({
                json: jest.fn().mockResolvedValueOnce(
                        'Unauthorized'
                )

            });
            global.fetch = mockFetch;

            const expectedError = new Error('Incorrect response format');

            // Act & Assert
            await expect(contentService.checkSpelling(text)).rejects.toThrow(expectedError);
        });
    });

    describe('getTranslationsForWords', () => {
        it('should return the same input when given an array of words', async () => {
            // Arrange
            const words = [{ title: 'hello' }, { title: 'world' }] as IWordToAnalyze[];
            const language = 'en-US';

            // Act
            const result = await contentService.getTranslationsForWords(words, language);

            // Assert
            expect(result).toEqual(words);
        });
    });
});