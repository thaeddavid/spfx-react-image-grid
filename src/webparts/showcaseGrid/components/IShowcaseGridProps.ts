export interface IShowcaseGridProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  gridItems: {
    imageUrl: string;
    title: string;
    description: string;
    linkUrl: string;
    linkText: string;
  }[];
}

export interface IShowcaseItem {
  imageUrl: string;
  title: string;
  description: string;
  linkUrl: string;
  linkText: string;
}
