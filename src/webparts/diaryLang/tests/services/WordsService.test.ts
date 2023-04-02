import { WordService } from '../../services/WordService';
import { IWordServices } from '../../services/interfaces/IWordServices';
import { ServiceKey } from '@microsoft/sp-core-library';

declare var global: any;

describe('WordService', () => {
    let wordService: IWordServices;

    beforeAll(async () => {
        global.Office = {};
        global.Word = {};
        wordService = new WordService();
        await wordService.initializeWord();
    });

    describe('getWordsFromHtml', () => {
        beforeAll(async () => {
            global.Office = {};
            global.Word = {};
        });

        it('should extract incorrect words from an HTML string', () => {
            const htmlString = '<div class="NormalTextRun">This is a test.</div><div class="SpellingErrorV2Themed">testt</div>';
            const result = wordService['getWordsFromHtml'](htmlString);
            expect(result).toHaveProperty('incorrectWords');
            expect(result.incorrectWords).toContain('testt');
            expect(result).toHaveProperty('noOfWords');
            expect(result.noOfWords).toBe(4);
        });
    });

    describe('checkIfIsInitiated', () => {
      it('should throw an error if the Office API is not initialized', () => {
        wordService['isOfficeInitialized'] = false;
        expect(() => wordService.checkIfIsInitiated()).toThrow('Office API is not initialized or Word object is not available');
      });
    });
});