import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import DiaryLang from './components/DiaryLang';
import { IDiaryLangProps } from './components/IDiaryLangProps';
import { WordService } from './services/WordService';
import { IWordServices } from './services/interfaces/IWordServices';
import { GptContentService, IContentService, ISpService, SpService } from './services';
import { Constants } from './constants/Constants';
import { initializeIcons } from 'office-ui-fabric-react';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IDiaryLangWebPartProps {
}

export default class DiaryLangWebPart extends BaseClientSideWebPart<IDiaryLangWebPartProps> {
  private wordService: IWordServices;
  private spService: ISpService;
  private contentService: IContentService;

  protected async onInit(): Promise<void> {
    initializeIcons();
    
    await SPComponentLoader.loadScript('https://appsforoffice.microsoft.com/lib/1/hosted/office.js', { globalExportsName: 'Office' });

    this.contentService = this.context.serviceScope.consume(GptContentService.serviceKey);
    this.wordService = this.context.serviceScope.consume(WordService.serviceKey);
    this.spService = this.context.serviceScope.consume(SpService.serviceKey);
    this.wordService.initializeWord();
    this.spService.setWebUrl(Constants.WEB_URL);
    await super.onInit();

  }

  public render(): void {
    const element: React.ReactElement<IDiaryLangProps> = React.createElement(
      DiaryLang,
      {
        wordService: this.wordService,
        spService: this.spService,
        contentService: this.contentService
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [

      ]
    };
  }
}
