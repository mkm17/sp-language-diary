import { IDropdownOption } from 'office-ui-fabric-react';
import { ILanguageItem, IWordToAnalyze } from '../models';
import { IContentService, ISpService, IWordServices } from '../services';

export interface IDiaryLangProps {
  wordService: IWordServices;
  spService: ISpService;
  contentService: IContentService;
}

export interface IDiaryLangState {
  selectedWords: IWordToAnalyze[];
  languages: IDropdownOption[];
  selectedLanguage: IDropdownOption;
  isLoading: boolean;
  showConfirmation: boolean;
  allCurrentSystemWords: ILanguageItem[];
}
