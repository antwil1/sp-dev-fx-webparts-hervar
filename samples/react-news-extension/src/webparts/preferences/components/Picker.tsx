import * as React from "react";
import SPService from "../../../services/SPService";
import {
  Chip,
  Group,
  Box,
  Modal,
  Button,
  Flex,
  Grid,
  useMantineTheme,
  ModalBaseStylesNames,
  Styles,
  Alert,
} from "@mantine/core";
import GraphService from "../../../services/GraphService";
import { ITerm } from "../types/Component.Types";
import { IconCheck, IconAlertTriangle } from "@tabler/icons-react";
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

  // Alla termer (för visning av namn)
  const [termsInfo, setTermsInfo] = React.useState<ITerm[]>([]);

  // Valda taggar = lista av GUID (string[])
  const [tags, setTags] = React.useState<string[]>([]);

  // Global state (ska vara string[] med ID:n)
  const [tagList, setTagList] = useRecoilState(tagsListAtom);

  const theme = useMantineTheme();
  const [loading, setLoading] = React.useState<boolean>(false);
  const [submitted, setSubmitted] = React.useState<boolean>(false);
  const [errorMsg, setErrorMsg] = React.useState<string>("");

  // Hämta alla termer + initiera valda taggar från global state
  React.useEffect(() => {
    async function fetchTaxonomy() {
      const terms = await SPService.getAllTermsByTermSet(termsetGuid);
      const termsResult: ITerm[] = (terms || []).map((t: any) => ({
        id: t.id,
        title: t.labels?.[0]?.name ?? "",
      }));
      setTermsInfo(termsResult);

      // tagList ska vara string[]
      setTags(Array.isArray(tagList) ? tagList : []);
    }
    fetchTaxonomy();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [termsetGuid]);

  // Hantera val i Chip.Group (vi får IDs direkt)
  const onTagChange = (selectedIds: string[]) => {
    setTags(selectedIds);
  };

  // Spara endast ID:n i Open Extension
  const onSavePreferences = async () => {
    setLoading(true);
    setSubmitted(false);
    setErrorMsg("");

    try {
      const extension = await GraphService.GetExtension(extensionName);

      const userSettings = {
        "@odata.type": "microsoft.graph.openTypeExtension",
        extensionName,
        Tags: tags, // endast ID:n
      };

      if (extension === null) {
        await GraphService.SavePreferences(userSettings);        // kastar vid fel
      } else {
        await GraphService.UpdatePreferences(userSettings, extensionName); // kastar vid fel (även 413)
      }

      // ✅ Lyckat
      setTagList(tags);
      CachingService.remove(`Preferences-${extensionName}-${loginName}`);
      CachingService.remove(`CuratedNews-UserPreferences-${loginName}`);
      window.dispatchEvent(new CustomEvent("curated:preferencesSaved", { detail: { extensionName, loginName } }));
      setSubmitted(true);
    } catch (err: any) {
      console.error("onSavePreferences error:", err);

      const status = err?.statusCode ?? err?.status;
      const rawMsg = (err?.message || "").toString().toLowerCase();

      let message = "Kunde inte spara dina inställningar.";

      // Grafens 2 KB-gräns – mappa till ett begripligt fel
      if (status === 413 || rawMsg.includes("maximum size supported for each extension is")) {
        message = "Du har valt för många taggar. Välj färre taggar och försök igen.";
      } else if (status === 401 || status === 403) {
        message = "Du har inte behörighet att spara dessa inställningar. Kontakta IT-avdelningen.";
      } else if (status === 404) {
        message = "Inställningslagringen kunde inte hittas. Kontrollera att extension-namnet är korrekt i webbdelens inställningar.";
      } else if (status >= 500) {
        message = "Det uppstod ett tillfälligt problem hos Microsoft Graph. Försök igen om en stund.";
      } else if (typeof err?.message === "string") {
        message = err.message;
      }

      setErrorMsg(message);
    } finally {
      setLoading(false);
    }
  };


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
          color:
            theme.colorScheme === "dark"
              ? theme.colors.dark[9]
              : theme.colors.gray[2],
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
                      <Chip
                        key={t.id}
                        value={t.id}
                        checked={isSelected}
                        variant="filled"
                      >
                        {t.title}
                      </Chip>
                    );
                  })}
              </Group>
            </Chip.Group>
          </Grid.Col>

          {/* FEL – röd, stängbar, ingen auto-close */}
          {errorMsg && (
            <Grid.Col span={12}>
              <Alert
                icon={<IconAlertTriangle size="1rem" />}
                title="Kunde inte spara"
                color="red"
                withCloseButton
                onClose={() => setErrorMsg("")}
              >
                {errorMsg}
              </Alert>
            </Grid.Col>
          )}

          {/* OK – grön, stängbar, ingen auto-close */}
          {submitted && !errorMsg && (
            <Grid.Col span={12}>
              <Alert
                icon={<IconCheck size="1rem" />}
                title="Klart!"
                color="green"
                withCloseButton
                onClose={() => setSubmitted(false)}
              >
                Dina inställningar har sparats.
              </Alert>
            </Grid.Col>
          )}

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
      </Modal>
    </div>
  );
};
