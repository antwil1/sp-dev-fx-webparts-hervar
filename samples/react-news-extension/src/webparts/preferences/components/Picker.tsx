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
  const [tagList, setTagList] = useRecoilState<string[]>(tagsListAtom as any);

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

  // Normalisera Graph-fel (statuskod & text)
  const getErrorInfo = (err: any) => {
    const status =
      err?.statusCode ??
      err?.status ??
      err?.responseStatus ??
      err?.code === "ResourceNotFound"
        ? 404
        : undefined;
    const message = (err?.message || "").toString();
    return { status, message: message.toLowerCase() };
  };

  // Spara endast ID:n i Open Extension (försök PATCH, om 404 → POST)
  const onSavePreferences = async () => {
    setLoading(true);
    setSubmitted(false);
    setErrorMsg("");

    const payload = {
      "@odata.type": "microsoft.graph.openTypeExtension",
      extensionName,
      Tags: tags, // endast ID:n
    };

    try {
      // 1) Försök hämta extension
      let exists = false;
      try {
        const ext = await GraphService.GetExtension(extensionName);
        exists = !!ext; // finns => PATCH
      } catch (getErr: any) {
        const { status } = getErrorInfo(getErr);
        if (status === 404) {
          exists = false; // skapa nytt med POST
        } else {
          throw getErr; // andra fel → hanteras nedan
        }
      }

      // 2) PATCH om det finns, annars POST
      if (exists) {
        await GraphService.UpdatePreferences(payload, extensionName);
      } else {
        await GraphService.SavePreferences(payload);
      }

      // ✅ Lyckat
      setTagList(tags);
      CachingService.remove(`Preferences-${extensionName}-${loginName}`);
      CachingService.remove(`CuratedNews-UserPreferences-${loginName}`);
      window.dispatchEvent(
        new CustomEvent("curated:preferencesSaved", {
          detail: { extensionName, loginName },
        })
      );
      setSubmitted(true);
    } catch (err: any) {
      console.error("onSavePreferences error:", err);

      const { status, message } = getErrorInfo(err);
      let uiMsg = "Kunde inte spara dina inställningar.";

      if (status === 413 || message.includes("maximum size supported for each extension is")) {
        uiMsg =
          "Du har valt för många taggar. Välj färre taggar och försök igen (Microsoft Graph har en 2 KB-gräns per extension).";
      } else if (status === 401 || status === 403) {
        uiMsg =
          "Du har inte behörighet att spara dessa inställningar. Kontakta IT-avdelningen.";
      } else if (status === 404) {
        uiMsg =
          "Inställningslagringen kunde inte hittas och kunde inte skapas. Kontrollera att extension-namnet är korrekt.";
      } else if (status && status >= 500) {
        uiMsg =
          "Det uppstod ett tillfälligt problem hos Microsoft Graph. Försök igen om en stund.";
      } else if (err?.message) {
        uiMsg = err.message;
      }

      setErrorMsg(uiMsg);
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
