import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ILibraryRatingItem } from '../types';

interface ISharePointItem {
  Id: number;
  Title: string;
  FileLeafRef: string;
  FileRef: string;
  AverageRating?: number;
  RatingCount?: number;
  Modified: string;
}

interface ISharePointItemsResponse {
  value: ISharePointItem[];
}

export class SharePointRatingsService {
  constructor(private readonly context: WebPartContext) {}

  public async getRatingItems(
    libraryTitle: string,
    pageSize: number,
    onlyCurrentUserRated: boolean
  ): Promise<ILibraryRatingItem[]> {
    const filter = onlyCurrentUserRated
      ? `&$filter=RatedBy/any(u:u/Id eq ${this.context.pageContext.legacyPageContext.userId})`
      : '';

    const endpoint =
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(
        libraryTitle
      )}')/items` +
      `?$select=Id,Title,FileLeafRef,FileRef,AverageRating,RatingCount,Modified` +
      `&$orderby=Modified desc&$top=${pageSize}${filter}`;

    const response = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    this.ensureSuccessful(response, '評価一覧の取得に失敗しました。');

    const payload = (await response.json()) as ISharePointItemsResponse;
    return payload.value.map((item) => ({
      id: item.Id,
      title: item.Title || item.FileLeafRef,
      fileLeafRef: item.FileLeafRef,
      serverRelativeUrl: item.FileRef,
      averageRating: Number(item.AverageRating || 0),
      ratingCount: Number(item.RatingCount || 0),
      modified: item.Modified
    }));
  }

  public async setRating(libraryTitle: string, itemId: number, rating: number): Promise<void> {
    const roundedRating = Math.max(0, Math.min(5, Math.round(rating)));
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(
      libraryTitle
    )}')/items(${itemId})/SetRating(rating=${roundedRating})`;

    const response = await this.context.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
      headers: {
        Accept: 'application/json;odata=nometadata'
      }
    });

    this.ensureSuccessful(response, '評価の更新に失敗しました。');
  }

  private ensureSuccessful(response: SPHttpClientResponse, message: string): void {
    if (!response.ok) {
      throw new Error(`${message} (HTTP ${response.status})`);
    }
  }
}
