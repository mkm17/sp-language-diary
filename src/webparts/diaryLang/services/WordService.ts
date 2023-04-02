import { IWordServices } from './interfaces/IWordServices';
import { ServiceKey } from '@microsoft/sp-core-library';

export class WordService implements IWordServices {
    public static readonly serviceKey: ServiceKey<IWordServices> = ServiceKey.create<IWordServices>('WordService', WordService);
    public Word: any;
    private isOfficeInitialized: boolean = false;
    private officeObj: typeof Office;

    /**
     * initializes the Word application
     */
    public async initializeWord(): Promise<void> {
        this.officeObj = Office;

        if (!this.officeObj) {
            console.warn('Office obj is not available.');
            return Promise.reject('Office obj is not available.');
        }

        try {
            this.officeObj.initialize = async () => {
                this.isOfficeInitialized = true;
            };
        } catch (error) {
            console.error('Error occurred during initialization:', error);
            return Promise.reject(error);
        }
    }

    /**
     * list of incorrect words in the current document
     * @returns list of incorrect words, and number of words
     * @deprecated
     */
    public async getListOfIncorrectWords(): Promise<{ incorrectWords: string[], noOfWords: number }> {
        this.checkIfIsInitiated();

        try {
            return Word.run(async (context) => {
                context.document.body.load('*');
                const documentHtml = context.document.body.getHtml();
                await context.sync();
                const { incorrectWords, noOfWords } = this.getWordsFromHtml(documentHtml.value);
                return { incorrectWords, noOfWords };
            });
        } catch (error) {
            console.error('Error occurred in getListOfIncorrectWords:', error);
            throw error;
        }
    }

    /**
     * returns the text in the current document
     * @returns the selected text in the current document
     */
    public async getText(): Promise<string> {
        this.checkIfIsInitiated();

        try {
            return Word.run(async (context) => {
                context.document.body.load('*');
                await context.sync();
                return context.document.body.text;
            });
        } catch (error) {
            console.error('Error occurred in getText:', error);
            throw error;
        }
    }

    /**
     * adds text at the end of the current document
     */
    public async addTextAtTheEndOfDocument(text: string): Promise<void> {
        this.checkIfIsInitiated();

        try {
            return Word.run(async (context) => {
                const docBody = context.document.body;
                const newParagraph = docBody.insertParagraph(text,
                    Word.InsertLocation.end);

                newParagraph.font.set({
                    name: 'Calibri',
                    size: 11,
                    italic: true,
                    color: 'C8C8C8'
                });

            });
        } catch (error) {
            console.error('Error occurred in addTextAtTheEnd:', error);
            throw error;
        }
    }

    /**
     * returns the selected text in the current document
     * @returns selected text
     */
    public async getSelectedText(): Promise<string> {
        this.checkIfIsInitiated();

        try {
            return Word.run(async (context) => {
                const selection = context.document.getSelection();

                selection.load('text');
                await context.sync();

                const selectedText = selection.text.trim();

                if (!selectedText) {
                    throw new Error('No text is selected');
                }

                return selectedText;
            });
        } catch (error) {
            console.error('Error occurred in getSelectedText:', error);
            throw error;
        }
    }

    /**
     * extract incorrect words from an HTML string
     * @param htmlString 
     * @returns list og incorrect Words and number of words
     * @deprecated
     */
    private getWordsFromHtml(htmlString: string): { incorrectWords: string[], noOfWords: number } {

        try {
            const div = document.createElement('div');
            div.innerHTML = htmlString;

            const totalWords = div.getElementsByClassName('NormalTextRun');
            let noOfWords = 0;

            for (let j = 0; j < totalWords.length; j++) {
                const words = totalWords[j];
                const splittedWords = words.innerHTML.split(' ');
                noOfWords += splittedWords.length;
            }

            const elements = div.getElementsByClassName('SpellingErrorV2Themed');

            const incorrectWords = [];

            for (let i = 0; i < elements.length; i++) {
                const element = elements[i];
                const textOfElement = element.innerHTML.trim().toLowerCase();
                incorrectWords.push(textOfElement);
            }

            return { incorrectWords: incorrectWords, noOfWords: noOfWords };
        } catch (error) {
            console.error('Error occurred in getWordsFromHtml:', error);
            throw error;
        }
    }

    /**
     * Check if the Office API is initialized and Word object is available
     */
    public checkIfIsInitiated(): void {
        if (!this.isOfficeInitialized || !Word) {
            throw new Error('Office API is not initialized or Word object is not available');
        }
    }

}