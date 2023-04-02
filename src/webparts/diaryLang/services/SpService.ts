import { ServiceKey } from '@microsoft/sp-core-library';
import { Web } from '@pnp/sp/presets/all';
import { IDropdownOption } from 'office-ui-fabric-react';
import { Constants } from '../constants/Constants';
import { ILanguageItem, IWordToAnalyze } from '../models';
import { ISpService } from './interfaces';

export class SpService implements ISpService {
    public static readonly serviceKey: ServiceKey<ISpService> = ServiceKey.create<ISpService>('SpService', SpService);
    private webUrl: string;

    /**
     * Sets the to webUrl property of the class to the specified @param web.
     * @param url 
     */
    public setWebUrl(url: string): void {
        this.webUrl = url;
    }


    /**
     *  Retrieves a list of language options from a SharePoint list 
     * @returns array of IDropdownOption objects
     */
    public async getLanguagesOptions(): Promise<IDropdownOption[]> {
        try {
            const languages = await Web(this.webUrl).lists.getByTitle(Constants.LANGUAGES_LIST_NAME).items.select('Id,Title').get();
            return languages.map(language => ({ key: language.Id, text: language.Title }));
        } catch (error) {
            console.error(`Error while getting language items`);
            throw error;
        }
    }

    /**
     * Retrieves an item from a SharePoint list called @param languageName where the Text property matches the word parameter.
     * @param language 
     * @param word 
     * @returns ILanguageItem object or null if no matching item was found
     */
    public async getItem(language: string, word: string): Promise<ILanguageItem> {
        try {
            const languageName = language ? language : Constants.BASE_LIST_NAME;
            const modifiedWord = word.trim().toLowerCase();
            const item = await Web(this.webUrl).lists.getByTitle(languageName).items
                .select('Id,Title,Text,Count')
                .filter(`Text eq '${modifiedWord}'`)
                .top(1)
                .get();
            return item ? item[0] : null;
        } catch (error) {
            console.error(`Error while getting item`);
            throw error;
        }
    }

    /**
     * Retrieves all items from a SharePoint list called @param languageName.
     * @param language 
     * @returns an array of ILanguageItem objects
     */
    public async getItems(language: string): Promise<ILanguageItem[]> {
        try {
            const languageName = language ? language : Constants.BASE_LIST_NAME;
            const items = await Web(this.webUrl).lists.getByTitle(languageName).items
                .select('Id,Title,Text,Count')
                .getAll();
            return items;
        } catch (error) {
            console.error(`Error while getting items`);
            throw error;
        }
    }

    /**
     *  Saves an array of IWordToAnalyze objects to a SharePoint list called languageName . 
     * If an itemId property is present on an item in the array, the corresponding SharePoint list item is updated with a new Count property value. 
     * If an itemId property is not present, a new SharePoint list item is created with the relevant property values.
     * @param analyzedItems 
     * @param language 
     */
    public async saveItems(analyzedItems: IWordToAnalyze[], language: string): Promise<void> {
        try {
            const languageName = language ? language : Constants.BASE_LIST_NAME;
            const web = Web(this.webUrl);
            const spList = web.lists.getByTitle(languageName);
            const batch = web.createBatch();

            for (const analyzedItem of analyzedItems) {

                if (analyzedItem.itemId) {
                    const count = analyzedItem.noOfOccurrences++;
                    spList.items.getById(analyzedItem.itemId).inBatch(batch).update({ Count: count, Skip: false });
                } else {
                    const { title, noOfOccurrences, translation } = analyzedItem;
                    const valueToUpdate = { Title: translation, Text: title, Count: noOfOccurrences, Skip: false };
                    spList.items.inBatch(batch).add(valueToUpdate);
                }
            }

            await batch.execute();
        } catch (error) {
            console.error(`Error while saving items`);
            throw error;
        }
    }

    /**
     * Updates a SharePoint list item with the specified itemId in a SharePoint list called @param language with a new Skip property value
     * @param itemId 
     * @param shouldSkip 
     * @param language 
     */
    public async updateItem(itemId: number, shouldSkip: boolean, language: string = Constants.BASE_LIST_NAME): Promise<void> {
        try {
            const list = Web(this.webUrl).lists.getByTitle(language);
            const item = await list.items.getById(itemId);
            await item.update({ 'Skip': shouldSkip });
        } catch (error) {
            console.error(`Error updating item with ID ${itemId} in ${language} list: ${error}`);
            throw error;
        }
    }

    /**
     * Adds a new item to a SharePoint report list with properties representing language statistics.
     * @param language 
     * @param totalWords 
     * @param incorrectWords 
     * @param newIncorrectWords 
     */
    public async setDiaryReportData(language: string, totalWords: number, incorrectWords: number, newIncorrectWords: number): Promise<void> {
        try {
            const list = Web(this.webUrl).lists.getByTitle(Constants.DIARY_LIST_NAME);
            await list.items.add({
                'Language': language,
                'TotalWords': totalWords,
                'IncorrectWords': incorrectWords,
                'NewIncorrectWords': newIncorrectWords
            });
        } catch (error) {
            console.error(`Error adding item to ${Constants.DIARY_LIST_NAME} list: ${error}`);
            throw error;
        }
    }
}