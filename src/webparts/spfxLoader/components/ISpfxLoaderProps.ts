import { MSGraphClient } from "@microsoft/sp-http";

export interface ISpfxLoaderProps {
  description: string;
  graphClient: MSGraphClient;
}
