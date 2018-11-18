import { PageContext } from "@microsoft/sp-page-context";

export interface IPageBannerProps {
  pageContext: PageContext;
  sourcePage: string;
  targetPage: string;
  modernizationCenterUrl: string;
  feedbackList: string;
  learnMoreUrl: string;
}
