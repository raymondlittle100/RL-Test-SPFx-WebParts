import { HttpClient } from '@microsoft/sp-http';

export interface ICreateModernPageProps {
  siteUrl: string,
  functionUrl: string,
  httpClient: HttpClient
}
