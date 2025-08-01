import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IModernSharePointDashboardProps {
  isDarkTheme: boolean;
  hasTeamsContext: boolean;
  context: WebPartContext;
}
