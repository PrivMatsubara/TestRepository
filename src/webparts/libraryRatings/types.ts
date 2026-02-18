export interface ILibraryRatingItem {
  id: number;
  title: string;
  fileLeafRef: string;
  serverRelativeUrl: string;
  averageRating: number;
  ratingCount: number;
  userRating?: number;
  modified: string;
}

export interface ILibraryRatingsProps {
  libraryTitle: string;
  pageSize: number;
  showOnlyCurrentUserRatings: boolean;
}
