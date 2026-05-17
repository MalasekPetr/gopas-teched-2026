import * as React from "react";
import { Panel, PanelType, TextField, Dropdown, PrimaryButton, DefaultButton, Stack } from "@fluentui/react";
import type { SPHttpClient } from "@microsoft/sp-http";
import { DefinitionService, type AlertDefinitionItem } from "../services/definitionService";

export interface IAlertPanelProps {
  siteUrl: string;
  listId: string;
  listTitle: string;
  currentUserEmail: string;
  spHttpClient: SPHttpClient;
  onClose: () => void;
}

const NOTIFY_OPTIONS = [
  { key: "All", text: "All changes" },
  { key: "New items", text: "New items only" },
  { key: "Modified", text: "Modified" },
  { key: "Deleted", text: "Deleted" },
];

export const AlertPanel: React.FC<IAlertPanelProps> = (props) => {
  const svc = React.useMemo(() => new DefinitionService(props.spHttpClient, props.siteUrl), [props.siteUrl]);
  const [email, setEmail] = React.useState(props.currentUserEmail);
  const [notifyOn, setNotifyOn] = React.useState<AlertDefinitionItem["NotifyOn"]>("All");
  const [existing, setExisting] = React.useState<AlertDefinitionItem | null>(null);

  React.useEffect(() => {
    svc.findForList(props.listId, props.currentUserEmail).then((found) => {
      if (!found) return;
      setExisting(found);
      setEmail(found.UserEmail);
      setNotifyOn(found.NotifyOn);
    });
  }, [props.listId]);

  const save = async (): Promise<void> => {
    if (existing) await svc.remove(existing.Id);
    await svc.create({
      Title: `${props.listTitle} – ${email}`,
      WatchedListId: props.listId,
      WatchedSiteUrl: props.siteUrl,
      UserEmail: email,
      NotifyOn: notifyOn,
      IsActive: true,
    });
    props.onClose();
  };

  const remove = async (): Promise<void> => {
    if (existing) await svc.remove(existing.Id);
    props.onClose();
  };

  return (
    <Panel isOpen headerText="Set Alert" type={PanelType.medium} onDismiss={props.onClose}>
      <Stack tokens={{ childrenGap: 12 }}>
        <TextField label="List" value={props.listTitle} readOnly />
        <Dropdown
          label="Notify me about"
          selectedKey={notifyOn}
          options={NOTIFY_OPTIONS}
          onChange={(_, o) => o && setNotifyOn(o.key as AlertDefinitionItem["NotifyOn"])}
        />
        <TextField label="Your email" value={email} onChange={(_, v) => setEmail(v || "")} />
        <Stack horizontal tokens={{ childrenGap: 8 }}>
          <PrimaryButton text="Save Alert" onClick={save} />
          {existing && <DefaultButton text="Remove Alert" onClick={remove} />}
          <DefaultButton text="Cancel" onClick={props.onClose} />
        </Stack>
      </Stack>
    </Panel>
  );
};
