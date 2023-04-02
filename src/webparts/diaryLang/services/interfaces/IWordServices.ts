export interface IWordServices {
    initializeWord(): void;
    getListOfIncorrectWords(): Promise<{ incorrectWords: string[], noOfWords: number }>;
    getSelectedText(): Promise<string>;
    addTextAtTheEndOfDocument(text: string): Promise<void>;
    checkIfIsInitiated(): void;
    getText(): Promise<string>;
}