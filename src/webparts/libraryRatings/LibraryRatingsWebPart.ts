import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneSlider, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { LibraryRatings } from './components/LibraryRatings';
import { ILibraryRatingsProps } from './types';
import { SharePointRatingsService } from './services/SharePointRatingsService';

export default class LibraryRatingsWebPart extends BaseClientSideWebPart<ILibraryRatingsProps> {
  private ratingsService!: SharePointRatingsService;

  protected onInit(): Promise<void> {
    this.ratingsService = new SharePointRatingsService(this.context);
    return super.onInit();
  }

  public render(): void {
    const element = React.createElement(LibraryRatings, {
      libraryTitle: this.properties.libraryTitle,
      pageSize: this.properties.pageSize || 20,
      showOnlyCurrentUserRatings: this.properties.showOnlyCurrentUserRatings || false,
      ratingsService: this.ratingsService
    });

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
        {
          header: { description: 'ライブラリ評価Webパーツ設定' },
          groups: [
            {
              groupName: '表示設定',
              groupFields: [
                PropertyPaneTextField('libraryTitle', {
                  label: '対象ドキュメントライブラリ名'
                }),
                PropertyPaneSlider('pageSize', {
                  label: '取得件数',
                  min: 5,
                  max: 100,
                  step: 5,
                  value: 20,
                  showValue: true
                }),
                PropertyPaneToggle('showOnlyCurrentUserRatings', {
                  label: '自分が評価したアイテムのみ表示'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
