import * as React from "react";
import SPService from "../../../services/SPService";
import {
  Chip, Group, Box, Modal, Button, Flex, Grid,
  useMantineTheme, ModalBaseStylesNames, Styles, Alert,
} from "@mantine/core";
import GraphService from "../../../services/GraphService";
import { ITerm } from "../types/Component.Types";
import { IconCheck } from "@tabler/icons-react";
import { useRecoilState } from "recoil";
import { tagsListAtom } from "../../../stores/appstore";
import CachingService from "../../../services/CachingService";

export interface IPickerProps {
  extensionName: string;
  termsetGuid: string;
  opened: boolean;
  close: () => void;
  loginName: string;
}

export const Picker: React.FC<IPickerProps> = (props) => {
  const { extensionName, termsetGuid, opened, close, loginName } = props;

  const theme = useMantineTheme();
  const [termsInfo, setTermsInfo] = React.useState<ITerm[]>([]);
  // ⬇️ Endast ID:n
  const [tags, setTags] = React.useState<string[]>([]);
  const [tagList, setTagList] = useRecoilState<string[]>(tagsListAtom);
  const [loading, setLoading] = React.useState<boolean>(false);
  const [submitted, setSubmitted] = React.useState<boolean>(false);

  React.useEffect(() => {
    async function fetchTaxonomy() {
      const terms = await SPService.getAllTermsByTermSet(termsetGuid);
      const termsResult: ITerm[] = (terms || []).map((t: any) => ({
        id: t.id,
        title: t.labels?.[0]?.name ?? t.name ?? t.id,
      }));
      setTermsInfo(termsResult);
      setTags(tagList); // atomen är redan string[]
    }
    fetchTaxonomy();
  }, []);

  const onSavePreferences = async () => {
    setLoading(true);
    setSubmitted(false);

    const extension = await GraphService.GetExtension(extensionName);

    // ⬇️ Skicka bara ID:n till Graph
    const userSettings = {
      "@odata.type": "microsoft.graph.openTypeExtension",
      extensionName: extensionName,
      Tags: tags as string[],
    };

    if (extension === null) {
      await GraphService.SavePreferences(userSettings);
    } else {
      await GraphService.UpdatePreferences(userSettings, extensionName);
    }

    // 1) Uppdatera lokalt UI
    setTagList(tags);
    setSubmitted(true);

    // 2) Rensa cache
    CachingService.remove(`Preferences-${extensionName}-${loginName}`);
    CachingService.remove(`CuratedNews-UserPreferences-${loginName}`);

    // 3) Signalera
    window.dispatchEvent(new CustomEvent("curated:preferencesSaved", {
      detail: { extensionName, loginName }
    }));

    setLoading(false);
  };

  const onTagChange = (selectedIds: string[]) => setTags([...selectedIds]);

  const modelHeaderStyles: Styles<ModalBaseStylesNames> = {
    header: {
      backgroundColor: "#d1d2d3ba",
      h2: { fontSize: "1.1rem" },
    },
  };

  return (
    <div>
      <Modal
        styles={modelHeaderStyles}
        size="lg"
        opened={opened}
        onClose={close}
        title="Uppdatera preferenser"
        centered
        overlayProps={{
          color: theme.colorScheme === "dark" ? theme.colors.dark[9] : theme.colors.gray[2],
          opacity: 0.55,
          blur: 3,
        }}
      >
        <Grid>
          <Grid.Col span={12}>
            <Chip.Group multiple value={tags} onChange={onTagChange}>
              <Group position="center" mt="md">
                {termsInfo.length > 0 &&
                  termsInfo.map((t: ITerm) => {
                    const isSelected = tags.includes(t.id);
                    return (
                      <Chip checked={isSelected} variant="filled" key={t.id} value={t.id}>
                        {t.title}
                      </Chip>
                    );
                  })}
              </Group>
            </Chip.Group>
          </Grid.Col>

          {!submitted && (
            <Grid.Col span={12}>
              <Flex gap="md" justify="flex-end">
                <Box w={200}>
                  <Button
                    loading={loading}
                    loaderPosition="left"
                    fullWidth
                    variant="gradient"
                    onClick={onSavePreferences}
                  >
                    Spara
                  </Button>
                </Box>
              </Flex>
            </Grid.Col>
          )}
        </Grid>

        {submitted && (
          <Alert
            icon={<IconCheck size="1rem" />}
            title="Klart!"
            color="green"
            withCloseButton
            onClose={() => setSubmitted(false)}
          >
            Dina inställningar har sparats. Allt är redo!
          </Alert>
        )}
      </Modal>
    </div>
  );
};
