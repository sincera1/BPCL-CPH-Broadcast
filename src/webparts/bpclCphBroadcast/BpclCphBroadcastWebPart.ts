import * as React from "react";
import * as ReactDOM from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";

import MyViewBroadcast from "./components/MyViewBroadcast";

import { SPFI, spfi, SPFx } from "@pnp/sp";

export interface IBroadcastWebPartProps {}

export default class BroadcastWebPart extends BaseClientSideWebPart<IBroadcastWebPartProps> {

  public render(): void {
    const sp: SPFI = spfi().using(SPFx(this.context));

    const element = React.createElement(MyViewBroadcast, {
      sp: sp,
      context: this.context,
    });

    ReactDOM.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}    