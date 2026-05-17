import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
} from "@microsoft/sp-listview-extensibility";
import * as React from "react";
import * as ReactDOM from "react-dom";
import { AlertPanel } from "./components/AlertPanel";

const LOG_SOURCE = "SpAlertDemoCommandSet";

export interface ISpAlertDemoProps {}

export default class SpAlertDemoCommandSet extends BaseListViewCommandSet<ISpAlertDemoProps> {
  private panelContainer: HTMLDivElement | null = null;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized");
    const cmd: Command = this.tryGetCommand("SET_ALERT");
    if (cmd) cmd.visible = true;
    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    if (event.itemId !== "SET_ALERT") return;

    // STEP A: Open the React panel — pass everything the panel needs from the
    // command set context (site URL, list id/title, current user email).
    if (!this.panelContainer) {
      this.panelContainer = document.body.appendChild(document.createElement("div"));
    }

    const pageCtx = this.context.pageContext;
    ReactDOM.render(
      React.createElement(AlertPanel, {
        siteUrl: pageCtx.web.absoluteUrl,
        listId: pageCtx.list!.id.toString(),
        listTitle: pageCtx.list!.title,
        currentUserEmail: pageCtx.user.email,
        spHttpClient: this.context.spHttpClient,
        onClose: () => ReactDOM.unmountComponentAtNode(this.panelContainer!),
      }),
      this.panelContainer
    );
  }
}
