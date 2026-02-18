import * as React from 'react';
import { PrimaryButton, Spinner, Stack, Text, Link, MessageBar, MessageBarType } from '@fluentui/react';
import { Rating } from '@fluentui/react/lib/Rating';
import { ILibraryRatingItem } from '../types';
import { SharePointRatingsService } from '../services/SharePointRatingsService';

export interface ILibraryRatingsComponentProps {
  libraryTitle: string;
  pageSize: number;
  showOnlyCurrentUserRatings: boolean;
  ratingsService: SharePointRatingsService;
}

export const LibraryRatings: React.FC<ILibraryRatingsComponentProps> = ({
  libraryTitle,
  pageSize,
  showOnlyCurrentUserRatings,
  ratingsService
}) => {
  const [items, setItems] = React.useState<ILibraryRatingItem[]>([]);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [savingItemId, setSavingItemId] = React.useState<number | undefined>(undefined);
  const [errorMessage, setErrorMessage] = React.useState<string | undefined>(undefined);

  const loadItems = React.useCallback(async () => {
    setLoading(true);
    setErrorMessage(undefined);
    try {
      const result = await ratingsService.getRatingItems(libraryTitle, pageSize, showOnlyCurrentUserRatings);
      setItems(result);
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : '評価一覧の読み込み中にエラーが発生しました。');
    } finally {
      setLoading(false);
    }
  }, [libraryTitle, pageSize, ratingsService, showOnlyCurrentUserRatings]);

  React.useEffect(() => {
    void loadItems();
  }, [loadItems]);

  const onRate = async (itemId: number, newRating?: number): Promise<void> => {
    if (typeof newRating !== 'number') {
      return;
    }

    setSavingItemId(itemId);
    setErrorMessage(undefined);

    try {
      await ratingsService.setRating(libraryTitle, itemId, newRating);
      await loadItems();
    } catch (error) {
      setErrorMessage(error instanceof Error ? error.message : '評価の更新中にエラーが発生しました。');
    } finally {
      setSavingItemId(undefined);
    }
  };

  if (loading) {
    return <Spinner label="評価データを取得中..." />;
  }

  return (
    <Stack tokens={{ childrenGap: 16 }}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
        <Text variant="xLarge">ライブラリ評価</Text>
        <PrimaryButton text="再読み込み" onClick={() => void loadItems()} disabled={savingItemId !== undefined} />
      </Stack>

      {errorMessage && <MessageBar messageBarType={MessageBarType.error}>{errorMessage}</MessageBar>}

      {items.length === 0 && <Text>評価対象アイテムが見つかりませんでした。</Text>}

      {items.map((item) => (
        <Stack key={item.id} tokens={{ childrenGap: 4 }}>
          <Link href={item.serverRelativeUrl} target="_blank">
            {item.title}
          </Link>
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center">
            <Rating
              max={5}
              rating={item.averageRating}
              onChange={(_, value) => void onRate(item.id, value)}
              allowZeroStars
              disabled={savingItemId === item.id}
            />
            <Text>
              平均: {item.averageRating.toFixed(1)} / 5（{item.ratingCount} 件）
            </Text>
          </Stack>
          <Text variant="small">更新日時: {new Date(item.modified).toLocaleString('ja-JP')}</Text>
        </Stack>
      ))}
    </Stack>
  );
};
