import { IFilePickerResult } from '@pnp/spfx-property-controls/lib/PropertyFieldFilePicker';

export interface IShowcaseGridProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  gridItems: IShowcaseItem[];
  columns: number;
  rows: number;
}

export interface IShowcaseItem {
  imageUrl: IFilePickerResult;
  title: string;
  description: string;
  linkUrl: string;
  linkText: string;
}
