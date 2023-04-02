export interface IWordToAnalyze {
    itemId?: number;
    title: string;
    isChecked?: boolean;
    noOfOccurrences?: number;
    translation?: string;
    suggestions?: string[];
}