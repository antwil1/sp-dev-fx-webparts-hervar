import * as React from "react";
import styles from "./Preferences.module.scss";
import { IPreferencesProps } from "./IPreferencesProps";
import { Container, Group, createStyles, ActionIcon, Card } from "@mantine/core";
import { IconSettings } from "@tabler/icons-react";
import GraphService from "../../../services/GraphService";
import { Picker } from "./Picker";
import SPService from "../../../services/SPService";
import { ITerm } from "../types/Component.Types";
import { useRecoilState } from "recoil";
import { tagsListAtom } from "../../../stores/appstore";
import CachingService from "../../../services/CachingService";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";

const useStyles = createStyles(() => ({
  tagWrapper: {
    display: "flex",
    flexDirection: "unset",
    margin: "3px 1em 3px 0",
    alignItems: "center",
    flexWrap: "wrap",
    justifyContent: "center",
    gap: "1rem",
  },
  tag: {
    color: "black",
    backgroundColor: "#EDEBE9",
    padding: "0 0.625rem",
    border: "0.0625rem solid transparent",
    borderRadius: "4px",
    height: "1.75rem",
    fontSize: "0.875rem",
    lineHeight: "calc(1.625rem)",
    whiteSpace: "nowrap",
    transition: "background-color 100ms ease 0s",
  },
}));

export const Preferences: React.FC<IPreferencesProps> = (props) => {
  const { extensionName, termsetGuid, loginName, title, context, enableCaching } = props;
  const { classes } = useStyles();

  // ⬇️ Nu är listan string[] (bara ID:n)
  const [tagList, setTagList] = useRecoilState<string[]>(tagsListAtom);

  // För att slå upp titlar för visning
  const [termsInfo, setTermsInfo] = React.useState<ITerm[]>([]);

  const dataCacheKey = `Preferences-${extensionName}-${loginName}`;
  const termsCacheKey = `Preferences-taxonomy-${termsetGuid}`;

  const onConfigure = () => context.propertyPane.open();

  const loadTerms = React.useCallback(async () => {
    const cached = CachingService.get<ITerm[]>(termsCacheKey);
    if (cached) { setTermsInfo(cached); return; }
    const terms = await SPService.getAllTermsByTermSet(termsetGuid);
    const mapped: ITerm[] = (terms || []).map((t: any) => ({
      id: t.id,
      title: t.labels?.[0]?.name ?? t.name ?? t.id,
    }));
    setTermsInfo(mapped);
    CachingService.set(termsCacheKey, mapped);
  }, [termsetGuid, termsCacheKey]);

  const getUserPreferences = async (): Promise<string[]> => {
    const cachedData = CachingService.get<string[]>(dataCacheKey);
    if (cachedData !== null) return cachedData;

    const result = await GraphService.GetPreferences(extensionName);
    const ids: string[] = (result && Array.isArray(result.Tags)) ? result.Tags : [];
    if (enableCaching) CachingService.set(dataCacheKey, ids);
    return ids;
  };

  const getPreferences = React.useCallback(async () => getUserPreferences(), []);

  React.useEffect(() => { loadTerms(); }, [loadTerms]);

  React.useEffect(() => {
    getPreferences().then((ids) => setTagList(ids)).catch(console.log);
  }, [getPreferences]);

  const [isPanelOpen, setIsPanelOpen] = React.useState<boolean>(false);
  const onViewPanelClick = (): void => setIsPanelOpen(true);
  const onViewPanelDismiss = (): void => setIsPanelOpen(false);

  if (!extensionName || !termsetGuid) {
    return (
      <Placeholder
        iconName="Edit"
        iconText="Configure your web part"
        description="Please provide the Microsoft Graph open extension name and term set Id."
        buttonLabel="Configure"
        onConfigure={onConfigure}
      />
    );
  }

  const titleFor = React.useCallback(
    (id: string) => termsInfo.find((t) => t.id === id)?.title ?? id,
    [termsInfo]
  );

  return (
    <Container>
      <Card withBorder shadow="sm" radius="md">
        <Card.Section withBorder inheritPadding py="xs">
          <Group position="apart">
            <h2 className={styles.sectionTitle}>{title}</h2>
            <ActionIcon onClick={onViewPanelClick} variant="outline" color="indigo">
              <IconSettings size="1rem" />
            </ActionIcon>
          </Group>
        </Card.Section>

        <Group position="apart" mt="md" mb="xs">
          {isPanelOpen && (
            <Picker
              extensionName={extensionName}
              termsetGuid={termsetGuid}
              opened={isPanelOpen}
              close={onViewPanelDismiss}
              loginName={loginName}
            />
          )}

          <div className={classes.tagWrapper}>
            {tagList.length > 0 &&
              tagList.map((id) => (
                <div className={classes.tag} key={id}>
                  {titleFor(id)}
                </div>
              ))}
          </div>
        </Group>
      </Card>
    </Container>
  );
};
