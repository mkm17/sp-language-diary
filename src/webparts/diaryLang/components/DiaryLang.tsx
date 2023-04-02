import { find } from 'lodash';
import {
  ActionButton, Checkbox, ComboBox, Dropdown, IComboBox,
  IComboBoxOption, IDropdownOption, Label, Spinner, SpinnerSize, TextField
} from 'office-ui-fabric-react';
import * as React from 'react';
import { ILanguageItem, IWordToAnalyze } from '../models';
import styles from './DiaryLang.module.scss';
import './DiaryLang.scss';
import { IDiaryLangProps, IDiaryLangState } from './IDiaryLangProps';

export default class DiaryLang extends React.Component<IDiaryLangProps, IDiaryLangState> {

  constructor(props: IDiaryLangProps) {
    super(props);
    this.state = {
      selectedWords: [],
      languages: [],
      selectedLanguage: null,
      allCurrentSystemWords: [],
      isLoading: false,
      showConfirmation: false
    };
  }

  public componentDidMount(): void {
    this.initData();
  }

  private async initData(): Promise<void> {
    const { spService } = this.props;
    const languages = await spService.getLanguagesOptions();
    this.setState({ languages });
  }

  public render(): React.ReactElement<IDiaryLangProps> {
    return (
      <div className={styles.diaryLang}>
        {this.renderLoader()}
        {this.renderActionButtons()}
        {this.renderSelectedWords()}
      </div>
    );
  }

  private renderActionButtons = () => {
    const { languages, selectedLanguage } = this.state;

    return <>
      <Dropdown
        options={languages}
        selectedKey={selectedLanguage ? selectedLanguage.key : null}
        onChange={this.onLanguageChange}
      />
      <ActionButton
        className={styles.textFunctionButton}
        iconProps={{ iconName: 'TextRotateHorizontal' }}
        onClick={this.checkText}
        disabled={!selectedLanguage}
      >
        Check current text
      </ActionButton>
      <ActionButton
        className={styles.textFunctionButton}
        iconProps={{ iconName: 'TextRotateHorizontal' }}
        onClick={this.getSelectedText}
        disabled={!selectedLanguage}
      >
        Get selected text
      </ActionButton></>;
  }

  private renderSelectedWords = () => {
    const { selectedWords } = this.state;

    return selectedWords && selectedWords.length > 0 && (
      <>
        <Label className={styles.incorrectWordsLabel}>Words</Label>
        <div className={styles.wordsToAnalyze}>
          {selectedWords.map((word, index) => (
            <div key={index}>{this.renderWord(word, index)}</div>
          ))}
        </div>
        <ActionButton className={styles.textFunctionButton} onClick={this.saveItems}>
          Save Items
        </ActionButton>
      </>
    );
  }

  private renderWord = (word: IWordToAnalyze, indexKey: number) => {
    const suggestionsOptions = word && word.suggestions ? word.suggestions.map((item) => ({ key: item, text: item })) : [];

    return (
      <div key={indexKey} className={styles.wordToAnalyze}>
        <div className={styles.halfOfWidth}>
          <Checkbox className={styles.checkboxIcon} checked={word.isChecked} onChange={() => {
            this.toggleWord(word);
          }} />
          {suggestionsOptions && suggestionsOptions.length > 0 ? (
            <ComboBox
              className={styles.combobox}
              allowFreeform={true}
              placeholder={word.title}
              defaultValue={word.title}
              text={word.title}
              onChange={(_event: React.FormEvent<IComboBox>, option?: IComboBoxOption, _index?: number, value?: string) => {
                const text = value ? value : option.text;
                this.wordTextChange(word, text);
              }}
              options={suggestionsOptions}
            />
          ) : (
            <TextField
              className={styles.combobox}
              deferredValidationTime={2000}
              defaultValue={word.title}
              onChange={(_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string) => {
                this.wordTextChange(word, newValue);
              }}
              onBlur={() => this.updateWord(word)}
            />
          )}
        </div>
        <div className={styles.halfOfWidth}>
          <span>{word.translation}</span>
          <span className={styles.numberOfWords}>{word.noOfOccurrences}</span>
        </div>
      </div>
    );
  }

  private renderLoader = () => {
    const { isLoading } = this.state;

    return isLoading && <div className={styles.loadingContainer}>
      <Spinner className={styles.loadingSpinner} size={SpinnerSize.large} />
    </div>;
  }

  private onLanguageChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    this.setState({ selectedLanguage: item });
  }

  private toggleWord = (word: IWordToAnalyze) => {
    const { selectedWords } = this.state;
    const updatedWords = selectedWords.map((item) => {
      if (item === word) {
        return { ...word, isChecked: !word.isChecked };
      }
      return item;
    });
    this.setState({ selectedWords: updatedWords });
  }

  private wordTextChange = (word: IWordToAnalyze, newValue: string) => {
    word.title = newValue;
    this.updateWord(word);
  }

  private updateWord = async (word: IWordToAnalyze) => {
    this.setState({ isLoading: true });

    const { selectedWords } = this.state;
    const { spService } = this.props;

    const itemFromDatabase = await spService.getItem(this.getCurrentLanguageSelection(), word.title);
    if (itemFromDatabase) {
      const updatedWords = selectedWords.map((item) => {
        if (item.title === word.title) {
          return { ...word, noOfOccurrences: itemFromDatabase.Count };
        }
        return item;
      });
      this.setState({ selectedWords: updatedWords });
    }

    this.setState({ isLoading: false });
  }

  private saveItems = async () => {
    this.setState({ isLoading: true });
    const { selectedWords, allCurrentSystemWords } = this.state;
    const { spService, contentService } = this.props;

    const currentLanguageName = this.getCurrentLanguageSelection();
    let getSelectedWordsOnly = selectedWords.filter((item) => item.isChecked);
    const systemWordsMap = this.createMapOfSystemWords(allCurrentSystemWords);
    this.assignValuesFromSystemItemsToSelectedWords(getSelectedWordsOnly, systemWordsMap);

    const onlyNewWordsFromSelected = getSelectedWordsOnly.filter((item) => item.noOfOccurrences === 0);

    getSelectedWordsOnly = await contentService.getTranslationsForWords(getSelectedWordsOnly, currentLanguageName);

    await this.saveReportData(getSelectedWordsOnly.length, onlyNewWordsFromSelected.length, currentLanguageName);
    await spService.saveItems(getSelectedWordsOnly, currentLanguageName);

    this.setState({ selectedWords: [], isLoading: false, showConfirmation: true });
  }

  private saveReportData = async (incorrectItemsNo: number, newItemsNo: number, languageName: string) => {
    const { wordService, spService } = this.props;

    const text = await wordService.getText();
    const allWords = text.split(' ');

    await spService.setDiaryReportData(
      languageName,
      allWords.length,
      incorrectItemsNo,
      newItemsNo
    );
  }

  private checkText = async (): Promise<void> => {
    this.setState({ isLoading: true });
    const selectedWords = [...this.state.selectedWords];

    const textFromDocument = await this.props.wordService.getText();
    const currentLanguageName = this.getCurrentLanguageSelection();

    const textCheckResult = await this.props.contentService.checkSpelling(textFromDocument, currentLanguageName);
    const { incorrectWords, suggestedText } = textCheckResult;

    await this.props.wordService.addTextAtTheEndOfDocument(suggestedText);

    const allSystemWords = await this.props.spService.getItems(currentLanguageName);

    for (const word of incorrectWords) {

      const currentSystemWord = this.findSystemWordForWord(allSystemWords, word);
      this.assignValuesFromSystemWordToWord(currentSystemWord, word);

      word.isChecked = true;
      selectedWords.push(word);
    }

    this.setState({
      selectedWords,
      allCurrentSystemWords: allSystemWords,
      isLoading: false
    });
  }

  private getSelectedText = async (): Promise<void> => {
    this.setState({ isLoading: true });
    const { wordService, spService } = this.props;
    const selectedWords = [...this.state.selectedWords];

    const selectedText = await wordService.getSelectedText();
    const trimmedLowerText = selectedText.trim().toLowerCase();

    const currentSystemWord = await spService.getItem(this.getCurrentLanguageSelection(), selectedText);

    const newSelectedWord: IWordToAnalyze = { title: trimmedLowerText, isChecked: true, suggestions: [] };
    this.assignValuesFromSystemWordToWord(currentSystemWord, newSelectedWord);
    selectedWords.push(newSelectedWord);

    this.setState({ selectedWords, isLoading: false });
  }

  private createMapOfSystemWords = (allCurrentSystemWords: ILanguageItem[]) => {
    const currentSystemWordsMap = new Map<string, any>();

    for (const systemItem of allCurrentSystemWords) {
      const systemWord = systemItem.Text ? systemItem.Text.toLowerCase() : '';
      currentSystemWordsMap.set(systemWord, { count: systemItem.Count, itemId: systemItem.Id, title: systemItem.Text, translation: systemItem.Title });
    }
    return currentSystemWordsMap;
  }

  private assignValuesFromSystemItemsToSelectedWords = (words: IWordToAnalyze[], systemWords: Map<string, any>) => {
    words.forEach((word) => {
      const systemWord = systemWords.get(word.title.toLowerCase());
      if (systemWord) {
        word.noOfOccurrences = systemWord.count;
        word.itemId = systemWord.itemId;
        word.translation = systemWord.translation;
      }
    });
  }

  private findSystemWordForWord = (allCurrentSystemWords: ILanguageItem[], word: IWordToAnalyze) => {
    return find(allCurrentSystemWords, (item) => {
      return item.Text && (item.Text.toLowerCase() === word.title.toLowerCase());
    });
  }

  private assignValuesFromSystemWordToWord = (currentSystemWord: ILanguageItem, word: IWordToAnalyze) => {
    word.noOfOccurrences = currentSystemWord ? currentSystemWord.Count : 0;
    word.itemId = currentSystemWord ? currentSystemWord.Id : null;
    word.translation = currentSystemWord ? currentSystemWord.Title : null;
  }

  private getCurrentLanguageSelection = () => {
    const { selectedLanguage } = this.state;
    return selectedLanguage ? selectedLanguage.text : null;
  }
}
