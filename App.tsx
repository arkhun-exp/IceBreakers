import * as React from "react";
import {
  Button,
  Card,
  Input,
  MessageBar,
  MessageBarActions,
  MessageBarBody,
  Spinner,
  Text,
  Textarea,
  makeStyles,
  tokens,
  webLightTheme,
} from "@fluentui/react-components";
import {
  ArrowRightRegular,
  ArrowSyncRegular,
  CalendarLtrRegular,
  CheckmarkCircleRegular,
  HistoryRegular,
  HomeRegular,
  PeopleRegular,
  PlayRegular,
  WarningRegular,
} from "@fluentui/react-icons";
import type {
  DataTable,
  GeneratedComponentProps,
  ReadableTableRow,
  UxAgentDataApi,
  WritableTableRow,
  exp_icebreakerhistory,
  exp_speakers,
} from "./RuntimeTypes";

type SpeakerRow = ReadableTableRow<exp_speakers>;
type HistoryRow = ReadableTableRow<exp_icebreakerhistory>;

type UserSettingsRow = {
  uilanguageid?: number;
  localeid?: number;
  decimalsymbol?: string;
  numberseparator?: string;
  currencysymbol?: string;
  dateformatstring?: string;
  dateseparator?: string;
};

type PageLanguage = { code: string; name: string; isRtl: boolean };
type XrmApi = {
  Utility?: {
    getGlobalContext?: () => {
      userSettings?: { languageId?: number; userId?: string };
    };
  };
  Navigation?: {
    navigateTo: (input: { pageType: "entitylist"; entityName: string }) => Promise<unknown>;
  };
};

type LookupReference = { table: string; id: string };
type LookupNormalization =
  | { state: "unset" }
  | { state: "resolved"; reference: LookupReference }
  | { state: "needs-target"; id: string }
  | { state: "invalid" };

type DashboardSnapshot = {
  speakers: SpeakerRow[];
  planned: HistoryRow[];
  completed: HistoryRow[];
  latestSent: HistoryRow | null;
  completedCount: number;
  userSettings: UserSettingsRow | null;
};

type DashboardLoad = {
  snapshot: DashboardSnapshot;
  errors: PageErrorView[];
};

type DashboardState = DashboardLoad & { loading: boolean };

type PageErrorKind = "read" | "write" | "indeterminate" | "network" | "render";
type PageErrorView = {
  operation: string;
  message: string;
  fallbackKey: string;
  code?: string;
  kind: PageErrorKind;
  retry?: () => void;
};
type PageErrorTelemetry = {
  trackException?: (error: unknown, properties?: Record<string, unknown>) => void;
};
type ActionDataApi = {
  executeAction?: (request: { name: string; parameters?: Record<string, unknown> }) => Promise<{
    ok: boolean;
    indeterminate?: boolean;
    outputs?: Record<string, unknown>;
    error?: { code?: string; message?: string };
  }>;
};

const PAGE_NAME = "IceBreaker";
const PAGE_LANGUAGES: Record<number, PageLanguage> = {
  1036: { code: "fr-FR", name: "Français (France)", isRtl: false },
};
const DEFAULT_LANGUAGE: PageLanguage = { code: "fr-FR", name: "Français (France)", isRtl: false };

const TRANSLATIONS: Record<string, Record<string, string>> = {
  "fr-FR": {
    title: "IceBreaker",
    subtitle: "",
    navHome: "Accueil",
    navSpeakers: "Speakers",
    navHistory: "IceBreakerHistory",
    navLabel: "Navigation principale",
    activeSpeakers: "Speakers actifs",
    toDoSpeakers: "Speakers à faire",
    completedIcebreakers: "IceBreakers réalisés",
    eligible: "Speakers éligibles",
    selectionReady: "Prêt pour le prochain",
    wheelDescription: "Sélection aléatoire d'un speaker à faire",
    wheelEmpty: "Aucun speaker éligible",
    spinAction: "Lancer la sélection",
    spinning: "Sélection en cours",
    drawSchedule: "Planifié le lundi prochain à 09:00",
    drawSuccess: "Demande transmise au flux.",
    drawPendingStatus: "En cours — attribution en attente.",
    latestSentSpeaker: "Dernier speaker envoyé : {speaker} · Envoyé",
    wheelPendingLabel: "Tirage en cours",
    wheelPendingValue: "En attente…",
    wheelResultLabel: "Dernier speaker envoyé",
    noSentSpeaker: "Aucune demande au statut « Envoyé » pour le moment.",
    drawUnavailable: "Ajoutez un speaker actif au statut « A faire » pour lancer le tirage.",
    resetSpeakers: "Réinitialiser les speakers à « A faire »",
    resettingSpeakers: "Réinitialisation…",
    resetSuccess: "Les speakers actifs ont été réinitialisés à « A faire ».",
    resetNoop: "Tous les speakers actifs sont déjà à « A faire ».",
    recentTitle: "Derniers IceBreakers réalisés",
    speakerColumn: "Speaker",
    dateColumn: "Date",
    activityColumn: "Activité",
    noRecent: "Aucun IceBreaker réalisé pour le moment.",
    closeTitle: "À clôturer",
    closeDescription: "Renseignez l’activité et le commentaire, puis marquez l’IceBreaker comme fait.",
    activityLabel: "Activité",
    activityPlaceholder: "Ex. : présentation de l’équipe",
    commentLabel: "Commentaire",
    commentPlaceholder: "Ajoutez un commentaire (facultatif)",
    markDone: "Marquer comme fait",
    saving: "Enregistrement…",
    closeSuccess: "IceBreaker marqué comme fait.",
    noPlanned: "Aucun IceBreaker planifié à clôturer.",
    unresolvedSpeaker: "Le speaker associé n’a pas pu être résolu. La clôture est désactivée pour éviter une mise à jour incorrecte.",
    plannedAlertTitle: "Plusieurs IceBreakers sont planifiés",
    plannedAlertDescription: "Vérifiez les planifications avant de lancer un autre tirage.",
    plannedCount: "{count} planifications",
    openPlans: "Ouvrir les planifications",
    speakerUnavailable: "Speaker non résolu",
    noSpeakerName: "Speaker sans nom",
    loading: "Chargement des données…",
    refresh: "Actualiser",
    refreshLabel: "Actualiser les données",
    autoRefreshCountdown: "Actualisation automatique dans {seconds} s",
    settingsLoadFailed: "Les préférences de format n’ont pas pu être chargées. Le format du navigateur sera utilisé.",
    appBrand: "EXPERTIME",
    appName: "IceBreaker",
    sidebarCaption: "Moments d’équipe",
    emptyNumber: "0",
    "error.read": "Impossible de charger les données. Réessayez.",
    "error.write": "La modification n’a pas pu être enregistrée. Actualisez avant de réessayer.",
    "error.network": "La requête n’a pas abouti. Vous pouvez réessayer.",
    "error.indeterminate": "Impossible de confirmer la fin de l’opération. Actualisez avant toute nouvelle tentative.",
    "error.render": "Un problème empêche l’affichage de cette page. Actualisez la page.",
    "error.tryAgain": "Réessayer",
  },
};

const ERROR_COPY = TRANSLATIONS["fr-FR"];

function getXrm(): XrmApi | undefined {
  return (window as unknown as { Xrm?: XrmApi }).Xrm;
}

function resolvePageLanguage(): PageLanguage {
  try {
    const languageId = getXrm()?.Utility?.getGlobalContext?.().userSettings?.languageId || 1036;
    return PAGE_LANGUAGES[languageId] || DEFAULT_LANGUAGE;
  } catch {
    return DEFAULT_LANGUAGE;
  }
}

function translate(key: string): string {
  const language = resolvePageLanguage();
  return TRANSLATIONS[language.code]?.[key] || ERROR_COPY[key] || key;
}

function reportPageError(
  pageName: string,
  operation: string,
  code: unknown,
  message: unknown,
  kind: PageErrorKind,
  retry?: () => void,
  appInsights?: PageErrorTelemetry,
): PageErrorView {
  const fields: Record<string, unknown> = {};
  for (const path of ["code", "errorCode", "message", "error.code", "error.message"]) {
    try {
      let value: unknown = message;
      for (const part of path.split(".")) {
        value = value !== null && typeof value === "object" ? (value as Record<string, unknown>)[part] : undefined;
      }
      fields[path] = value;
    } catch {
      fields[path] = undefined;
    }
  }
  let normalizedCode = "";
  for (const candidate of [code, fields["code"], fields["errorCode"], fields["error.code"]]) {
    const text = typeof candidate === "string" ? candidate : typeof candidate === "number" && Number.isFinite(candidate) ? String(candidate) : "";
    normalizedCode = text.replace(/[\s\u0000-\u001f\u007f-\u009f]+/g, " ").trim().slice(0, 80);
    if (normalizedCode) break;
  }
  let normalizedMessage = "";
  for (const candidate of [message, fields["message"], fields["error.message"]]) {
    const line = typeof candidate === "string" ? candidate.trim().split(/[\r\n\u2028\u2029]/)[0].replace(/[\s\u0000-\u001f\u007f-\u009f]+/g, " ").trim() : "";
    if (!line || /^\[object [^\]]*\]$/.test(line) || /^at\s.*(?::\d+:\d+\)?|\((?:native|<anonymous>)\))$/.test(line)) continue;
    normalizedMessage = line.length > 300 ? line.slice(0, 299) + "…" : line;
    break;
  }
  console.error("[" + String(pageName).replace(/[\r\n\]]/g, "") + "] " + operation + " failed", { code: normalizedCode, message: normalizedMessage });
  const quiet = kind === "indeterminate" || kind === "render";
  const view: PageErrorView = { operation, message: quiet ? "" : normalizedMessage, fallbackKey: "error." + kind, kind };
  if (normalizedCode) view.code = normalizedCode;
  if (typeof retry === "function" && (kind === "read" || kind === "network")) view.retry = retry;
  return view;
}

function PageErrorBar(props: { error: PageErrorView | null }) {
  const { error } = props;
  if (!error) return null;
  return (
    <MessageBar intent="error">
      <MessageBarBody>{error.message || translate(error.fallbackKey)}</MessageBarBody>
      {error.retry && (
        <MessageBarActions>
          <Button appearance="secondary" onClick={error.retry}>{translate("error.tryAgain")}</Button>
        </MessageBarActions>
      )}
    </MessageBar>
  );
}

class PageErrorBoundary extends React.Component<
  { children: React.ReactNode; appInsights?: PageErrorTelemetry },
  { failed: boolean }
> {
  public state = { failed: false };

  public static getDerivedStateFromError(): { failed: boolean } {
    return { failed: true };
  }

  public componentDidCatch(error: Error): void {
    reportPageError(PAGE_NAME, "render", undefined, error, "render");
  }

  public render(): React.ReactNode {
    if (this.state.failed) {
      return (
        <MessageBar intent="error">
          <MessageBarBody>{translate("error.render")}</MessageBarBody>
        </MessageBar>
      );
    }
    return this.props.children;
  }
}

const useStyles = makeStyles({
  root: {
    height: "100%",
    overflowY: "auto",
    boxSizing: "border-box",
    display: "flex",
    backgroundColor: "#F7F6FC",
    color: tokens.colorNeutralForeground1,
    fontFamily: "Segoe UI, Arial, sans-serif",
    "@media (max-width: 760px)": { flexDirection: "column" },
  },
  sidebar: {
    width: "15rem",
    flexShrink: 0,
    padding: "1.5rem 1rem",
    backgroundColor: "#FFFFFF",
    borderRightColor: "#E6E2F3",
    borderRightStyle: "solid",
    borderRightWidth: "1px",
    display: "flex",
    flexDirection: "column",
    gap: "1.75rem",
    "@media (max-width: 760px)": {
      width: "100%",
      boxSizing: "border-box",
      padding: "0.75rem 1rem",
      borderRightWidth: "0",
      borderBottomColor: "#E6E2F3",
      borderBottomStyle: "solid",
      borderBottomWidth: "1px",
      gap: "0.5rem",
    },
  },
  brand: { display: "flex", alignItems: "center", gap: "0.75rem", paddingInlineStart: "0.35rem" },
  brandMark: {
    width: "2.6rem",
    height: "2.6rem",
    borderRadius: "0.85rem",
    backgroundColor: "#5B3FD3",
    color: "#FFFFFF",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    fontWeight: tokens.fontWeightBold,
    fontSize: tokens.fontSizeHero800,
    flexShrink: 0,
  },
  brandText: { display: "flex", flexDirection: "column", gap: "0.1rem" },
  brandKicker: { color: "#746A91", fontSize: tokens.fontSizeBase100, letterSpacing: "0.12em", fontWeight: tokens.fontWeightSemibold },
  brandName: { color: "#242424", fontSize: tokens.fontSizeBase400, fontWeight: tokens.fontWeightSemibold },
  nav: {
    display: "flex",
    flexDirection: "column",
    gap: "0.4rem",
    "@media (max-width: 760px)": { flexDirection: "row", overflowX: "auto" },
  },
  navLink: {
    width: "100%",
    justifyContent: "flex-start",
    borderRadius: "0.75rem",
    minHeight: "2.75rem",
    color: "#5E5A68",
    "&[aria-current='page']": { backgroundColor: "#F0ECFF", color: "#4B32B8", fontWeight: tokens.fontWeightSemibold },
    "@media (max-width: 760px)": { width: "auto", flexShrink: 0 },
    "&:hover": { backgroundColor: "#F3F0FB", color: "#4B32B8" },
  },
  sidebarFoot: {
    marginTop: "auto",
    padding: "0.75rem",
    color: "#746A91",
    fontSize: tokens.fontSizeBase200,
    "@media (max-width: 760px)": { display: "none" },
  },
  main: {
    flex: 1,
    minWidth: 0,
    display: "flex",
    flexDirection: "column",
    gap: "1.5rem",
    padding: "2rem 2.25rem 2.5rem",
    "@media (max-width: 760px)": { padding: "1.25rem 1rem 2rem", gap: "1rem" },
  },
  header: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: "1rem" },
  headingBlock: { display: "flex", flexDirection: "column", gap: "0.35rem" },
  pageTitle: { color: "#242424", fontSize: "1.8rem", lineHeight: 1.2, fontWeight: tokens.fontWeightBold },
  subtitle: { color: "#746A91", fontSize: tokens.fontSizeBase300 },
  refreshControls: { display: "flex", flexDirection: "column", alignItems: "flex-end", gap: "0.2rem", flexShrink: 0 },
  refreshButton: { color: "#5B3FD3" },
  refreshCountdown: { color: "#746A91", fontSize: tokens.fontSizeBase100 },
  kpiGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(3, minmax(0, 1fr))",
    gap: "1rem",
    "@media (max-width: 1020px)": { gridTemplateColumns: "repeat(2, minmax(0, 1fr))" },
    "@media (max-width: 620px)": { gridTemplateColumns: "1fr" },
  },
  kpiCard: {
    minWidth: 0,
    flexShrink: 0,
    padding: "1.25rem",
    borderRadius: "12px",
    backgroundColor: "#FFFFFF",
    boxShadow: "0 2px 10px rgba(44, 34, 86, 0.06)",
    borderTopColor: "#E6E2F3",
    borderTopStyle: "solid",
    borderTopWidth: "1px",
    borderRightColor: "#E6E2F3",
    borderRightStyle: "solid",
    borderRightWidth: "1px",
    borderBottomColor: "#E6E2F3",
    borderBottomStyle: "solid",
    borderBottomWidth: "1px",
    borderLeftColor: "#E6E2F3",
    borderLeftStyle: "solid",
    borderLeftWidth: "1px",
    display: "flex",
    flexDirection: "column",
    gap: "0.8rem",
  },
  kpiTop: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: "0.75rem" },
  kpiLabel: { color: "#666173", fontSize: tokens.fontSizeBase300, fontWeight: tokens.fontWeightMedium },
  kpiIcon: {
    width: "2.4rem",
    height: "2.4rem",
    borderRadius: "0.75rem",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    backgroundColor: "#F0ECFF",
    color: "#5B3FD3",
  },
  kpiValue: { color: "#242424", fontSize: "2rem", lineHeight: 1, fontWeight: tokens.fontWeightBold, fontVariantNumeric: "tabular-nums" },
  kpiFoot: { color: "#777282", fontSize: tokens.fontSizeBase200 },
  contentGrid: {
    display: "grid",
    gridTemplateColumns: "minmax(0, 1.35fr) minmax(18rem, 0.9fr)",
    gap: "1rem",
    alignItems: "stretch",
    "@media (max-width: 1040px)": { gridTemplateColumns: "1fr" },
  },
  panel: {
    minWidth: 0,
    flexShrink: 0,
    padding: "1.35rem",
    borderRadius: "12px",
    backgroundColor: "#FFFFFF",
    borderTopColor: "#E6E2F3",
    borderTopStyle: "solid",
    borderTopWidth: "1px",
    borderRightColor: "#E6E2F3",
    borderRightStyle: "solid",
    borderRightWidth: "1px",
    borderBottomColor: "#E6E2F3",
    borderBottomStyle: "solid",
    borderBottomWidth: "1px",
    borderLeftColor: "#E6E2F3",
    borderLeftStyle: "solid",
    borderLeftWidth: "1px",
    boxShadow: "0 2px 10px rgba(44, 34, 86, 0.05)",
    display: "flex",
    flexDirection: "column",
    gap: "1rem",
  },
  panelHeader: { display: "flex", flexDirection: "column", gap: "0.3rem" },
  sectionTitle: { color: "#242424", fontSize: tokens.fontSizeBase500, fontWeight: tokens.fontWeightSemibold },
  sectionDescription: { color: "#746A91", fontSize: tokens.fontSizeBase200, lineHeight: 1.5 },
  selectionLayout: {
    display: "grid",
    gridTemplateColumns: "minmax(12rem, 0.9fr) minmax(13rem, 1fr)",
    gap: "1.25rem",
    alignItems: "center",
    "@media (max-width: 560px)": { gridTemplateColumns: "1fr" },
  },
  wheelColumn: { display: "flex", flexDirection: "column", alignItems: "center", gap: "0.75rem", minWidth: 0 },
  wheelWrap: { position: "relative", width: "min(100%, 17rem)", aspectRatio: "1 / 1", display: "flex", alignItems: "center", justifyContent: "center" },
  wheel: {
    width: "100%",
    height: "100%",
    borderRadius: "50%",
    borderTopColor: "#FFFFFF",
    borderTopStyle: "solid",
    borderTopWidth: "7px",
    borderRightColor: "#FFFFFF",
    borderRightStyle: "solid",
    borderRightWidth: "7px",
    borderBottomColor: "#FFFFFF",
    borderBottomStyle: "solid",
    borderBottomWidth: "7px",
    borderLeftColor: "#FFFFFF",
    borderLeftStyle: "solid",
    borderLeftWidth: "7px",
    boxShadow: "0 5px 18px rgba(47, 33, 99, 0.16)",
  },
  wheelPointer: {
    position: "absolute",
    zIndex: 2,
    top: "-0.15rem",
    width: 0,
    height: 0,
    borderLeftWidth: "0.65rem",
    borderLeftStyle: "solid",
    borderLeftColor: "transparent",
    borderRightWidth: "0.65rem",
    borderRightStyle: "solid",
    borderRightColor: "transparent",
    borderTopWidth: 0,
    borderBottomWidth: "1.15rem",
    borderBottomStyle: "solid",
    borderBottomColor: "#5B3FD3",
    filter: "drop-shadow(0 2px 2px rgba(30, 20, 70, 0.2))",
  },
  wheelCenter: {
    position: "absolute",
    zIndex: 1,
    width: "43%",
    aspectRatio: "1 / 1",
    borderRadius: "50%",
    backgroundColor: "#FFFFFF",
    borderTopColor: "#E6E2F3",
    borderTopStyle: "solid",
    borderTopWidth: "1px",
    borderRightColor: "#E6E2F3",
    borderRightStyle: "solid",
    borderRightWidth: "1px",
    borderBottomColor: "#E6E2F3",
    borderBottomStyle: "solid",
    borderBottomWidth: "1px",
    borderLeftColor: "#E6E2F3",
    borderLeftStyle: "solid",
    borderLeftWidth: "1px",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
    padding: "0.5rem",
    textAlign: "center",
    boxShadow: "0 2px 8px rgba(47, 33, 99, 0.1)",
  },
  wheelCenterLabel: { color: "#746A91", fontSize: "0.68rem", lineHeight: 1.2 },
  wheelCenterName: { color: "#4B32B8", fontSize: tokens.fontSizeBase300, fontWeight: tokens.fontWeightSemibold, lineHeight: 1.25, overflowWrap: "anywhere" },
  eligibleCaption: { color: "#746A91", fontSize: tokens.fontSizeBase200, textAlign: "center" },
  eligibleNames: { display: "flex", flexWrap: "wrap", justifyContent: "center", gap: "0.4rem", maxWidth: "100%" },
  eligibleChip: { padding: "0.25rem 0.55rem", borderRadius: "999px", backgroundColor: "#F3F0FB", color: "#51456E", fontSize: tokens.fontSizeBase100 },
  actionColumn: { display: "flex", flexDirection: "column", alignItems: "flex-start", gap: "0.9rem" },
  actionEyebrow: { color: "#746A91", fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold },
  actionTitle: { color: "#242424", fontSize: "1.35rem", lineHeight: 1.25, fontWeight: tokens.fontWeightSemibold },
  primaryAction: { backgroundColor: "#5B3FD3", color: "#FFFFFF", borderRadius: "0.75rem", minHeight: "2.9rem", paddingInline: "1.15rem", "&:hover": { backgroundColor: "#4B32B8", color: "#FFFFFF" } },
  secondaryAction: { color: "#5B3FD3", borderRadius: "0.75rem", minHeight: "2.75rem" },
  drawStatus: { color: "#51456E", fontSize: tokens.fontSizeBase200, lineHeight: 1.5 },
  muted: { color: "#746A91", fontSize: tokens.fontSizeBase200, lineHeight: 1.5 },
  resultLine: { display: "flex", alignItems: "center", gap: "0.5rem", color: "#3B8A62", fontSize: tokens.fontSizeBase200 },
  tablePanel: { minWidth: 0, flexShrink: 0, padding: "1.35rem", borderRadius: "12px", backgroundColor: "#FFFFFF", borderTopColor: "#E6E2F3", borderTopStyle: "solid", borderTopWidth: "1px", borderRightColor: "#E6E2F3", borderRightStyle: "solid", borderRightWidth: "1px", borderBottomColor: "#E6E2F3", borderBottomStyle: "solid", borderBottomWidth: "1px", borderLeftColor: "#E6E2F3", borderLeftStyle: "solid", borderLeftWidth: "1px", boxShadow: "0 2px 10px rgba(44, 34, 86, 0.05)" },
  tableWrap: { width: "100%", overflowX: "auto" },
  table: { width: "100%", borderCollapse: "collapse", textAlign: "start" },
  tableHeadCell: { padding: "0.85rem 0.75rem", color: "#746A91", fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold, textAlign: "start", borderBottomColor: "#E6E2F3", borderBottomStyle: "solid", borderBottomWidth: "1px", whiteSpace: "nowrap" },
  tableCell: { padding: "0.9rem 0.75rem", color: "#35313E", fontSize: tokens.fontSizeBase300, borderBottomColor: "#F0EDF6", borderBottomStyle: "solid", borderBottomWidth: "1px", verticalAlign: "middle" },
  tableStrong: { fontWeight: tokens.fontWeightSemibold, color: "#242424" },
  statusRow: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) auto", gap: "0.75rem", alignItems: "center", padding: "0.75rem 0", borderBottomColor: "#F0EDF6", borderBottomStyle: "solid", borderBottomWidth: "1px" },
  statusName: { color: "#35313E", fontSize: tokens.fontSizeBase300, fontWeight: tokens.fontWeightMedium, overflowWrap: "anywhere" },
  statusDate: { color: "#746A91", fontSize: tokens.fontSizeBase200, whiteSpace: "nowrap" },
  completionFields: { display: "flex", flexDirection: "column", gap: "0.85rem" },
  fieldLabel: { color: "#514C5D", fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold },
  fieldGroup: { display: "flex", flexDirection: "column", gap: "0.35rem" },
  completionButton: { alignSelf: "flex-start", backgroundColor: "#5B3FD3", color: "#FFFFFF", borderRadius: "0.75rem", minHeight: "2.75rem", "&:hover": { backgroundColor: "#4B32B8", color: "#FFFFFF" } },
  alertPanel: { borderTopColor: "#F1B7B7", borderTopStyle: "solid", borderTopWidth: "1px", borderRightColor: "#F1B7B7", borderRightStyle: "solid", borderRightWidth: "1px", borderBottomColor: "#F1B7B7", borderBottomStyle: "solid", borderBottomWidth: "1px", borderLeftColor: "#F1B7B7", borderLeftStyle: "solid", borderLeftWidth: "1px", backgroundColor: "#FFF7F7" },
  alertHeading: { display: "flex", alignItems: "center", gap: "0.65rem", color: "#A4262C", fontSize: tokens.fontSizeBase400, fontWeight: tokens.fontWeightSemibold },
  alertCopy: { color: "#713438", fontSize: tokens.fontSizeBase200, lineHeight: 1.5 },
  alertCount: { color: "#A4262C", fontSize: tokens.fontSizeBase200, fontWeight: tokens.fontWeightSemibold },
  alertButton: {
    alignSelf: "flex-start",
    color: "#A4262C",
    borderTopColor: "#D13438",
    borderTopStyle: "solid",
    borderTopWidth: "1px",
    borderRightColor: "#D13438",
    borderRightStyle: "solid",
    borderRightWidth: "1px",
    borderBottomColor: "#D13438",
    borderBottomStyle: "solid",
    borderBottomWidth: "1px",
    borderLeftColor: "#D13438",
    borderLeftStyle: "solid",
    borderLeftWidth: "1px",
  },
  emptyState: { padding: "1.5rem 0.75rem", color: "#746A91", fontSize: tokens.fontSizeBase300, textAlign: "center" },
  inlineLoading: { display: "flex", alignItems: "center", gap: "0.6rem", color: "#746A91", fontSize: tokens.fontSizeBase200 },
  errorStack: { display: "flex", flexDirection: "column", gap: "0.5rem" },
});

const BRAND_THEME = {
  ...webLightTheme,
  colorBrandBackground: "#5B3FD3",
  colorBrandBackgroundHover: "#4B32B8",
  colorBrandForeground1: "#5B3FD3",
  colorNeutralBackground1: "#FFFFFF",
  colorNeutralBackground2: "#F7F6FC",
  colorNeutralForeground1: "#242424",
  borderRadiusMedium: "12px",
};

function themeToVars(theme: typeof BRAND_THEME): React.CSSProperties {
  const vars: Record<string, string> = {};
  Object.entries(theme).forEach(([key, value]) => { vars[`--${key}`] = String(value); });
  return vars as React.CSSProperties;
}

const winAny = window as unknown as Record<string, unknown>;
const CACHE_KEY = "__ppIceBreakerDashboard_snapshot";
const INFLIGHT_KEY = "__ppIceBreakerDashboard_inflight";
const USER_SETTINGS_SELECT: string[] = [
  "uilanguageid",
  "localeid",
  "decimalsymbol",
  "numberseparator",
  "currencysymbol",
  "dateformatstring",
  "dateseparator",
];

const LOOKUP_GUID = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
const LOOKUP_QUALIFIED = /^\/([a-z_][a-z0-9_]*)\(([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\)$/;

function normalizeLookupReference(
  row: unknown,
  property: string,
  allowedTargets: readonly string[],
): LookupNormalization {
  if (typeof row !== "object" || row === null || Array.isArray(row) || allowedTargets.length === 0) return { state: "invalid" };
  const record = row as Record<string, unknown>;
  if (!Object.prototype.hasOwnProperty.call(record, property)) return { state: "invalid" };
  const raw = record[property];
  if (raw === null) return { state: "unset" };
  if (typeof raw !== "string") return { state: "invalid" };
  const value = raw.trim();
  const qualified = LOOKUP_QUALIFIED.exec(value);
  if (!qualified && !LOOKUP_GUID.test(value)) return { state: "invalid" };
  const prefix = qualified ? qualified[1] : undefined;
  const id = (qualified ? qualified[2] : value).toLowerCase();
  if (prefix !== undefined && !allowedTargets.includes(prefix)) return { state: "invalid" };
  const annotationKey = `${property}@Microsoft.Dynamics.CRM.lookuplogicalname`;
  let annotation: string | undefined;
  if (Object.prototype.hasOwnProperty.call(record, annotationKey)) {
    const named = record[annotationKey];
    if (typeof named !== "string" || !allowedTargets.includes(named)) return { state: "invalid" };
    annotation = named;
  }
  if (annotation !== undefined && prefix !== undefined && annotation !== prefix) return { state: "invalid" };
  const table = annotation ?? prefix ?? (allowedTargets.length === 1 ? allowedTargets[0] : undefined);
  return table === undefined ? { state: "needs-target", id } : { state: "resolved", reference: { table, id } };
}

function getHistorySpeakerLabel(row: HistoryRow): string {
  const label = row["_exp_speaker_value@OData.Community.Display.V1.FormattedValue"];
  return typeof label === "string" && label.trim() ? label : translate("speakerUnavailable");
}

async function collectRows<T>(firstPage: DataTable<T>): Promise<T[]> {
  const rows = firstPage.rows.slice();
  let page = firstPage;
  while (page.hasMoreRows) {
    if (!page.loadMoreRows) throw new Error("Dataverse reported more rows but did not provide a continuation reader.");
    page = await page.loadMoreRows();
    rows.push(...page.rows);
  }
  return rows;
}

async function readAllSpeakers(dataApi: UxAgentDataApi): Promise<SpeakerRow[]> {
  const first = await dataApi.queryTable("exp_speakers", {
    select: ["exp_speakersid", "exp_name", "exp_actif", "exp_statut", "exp_derniererealisation"],
    filter: "exp_actif eq 1",
    orderBy: "exp_name asc",
    pageSize: 250,
  });
  return collectRows(first);
}

async function readAllPlanned(dataApi: UxAgentDataApi): Promise<HistoryRow[]> {
  const first = await dataApi.queryTable("exp_icebreakerhistory", {
    select: ["exp_icebreakerhistoryid", "exp_name", "exp_activite", "exp_commentaires", "exp_date", "exp_reference", "exp_statut", "_exp_speaker_value"],
    filter: "exp_statut eq 791170000",
    orderBy: "exp_date asc",
    pageSize: 250,
  });
  return collectRows(first);
}

async function readRecentCompleted(dataApi: UxAgentDataApi): Promise<HistoryRow[]> {
  const result = await dataApi.queryTable("exp_icebreakerhistory", {
    select: ["exp_icebreakerhistoryid", "exp_name", "exp_activite", "exp_date", "exp_statut", "_exp_speaker_value"],
    filter: "exp_statut eq 791170002",
    orderBy: "exp_date desc",
    pageSize: 250,
  });
  return collectRows(result);
}

async function readLatestSent(dataApi: UxAgentDataApi): Promise<HistoryRow | null> {
  const result = await dataApi.queryTable("exp_icebreakerhistory", {
    select: ["exp_icebreakerhistoryid", "exp_name", "exp_date", "exp_statut", "_exp_speaker_value"],
    filter: "exp_statut eq 791170003",
    orderBy: "createdon desc,exp_icebreakerhistoryid desc",
    pageSize: 1,
  });
  return result.rows[0] ?? null;
}

type UserSettingsReader = {
  retrieveRow: (tableName: "usersettings", options: { id: string; select: string[] }) => Promise<UserSettingsRow>;
};

async function readUserSettings(dataApi: UxAgentDataApi): Promise<UserSettingsRow | null> {
  const userId = currentUserId();
  if (!userId) return null;
  const cacheKey = `__ppIceBreaker_userSettings_${userId}`;
  const inflightKey = `__ppIceBreaker_userSettingsInflight_${userId}`;
  const cached = winAny[cacheKey] as UserSettingsRow | undefined;
  if (cached !== undefined) return cached;

  let pending = winAny[inflightKey] as Promise<UserSettingsRow> | undefined;
  if (!pending) {
    const settingsApi = dataApi as unknown as UserSettingsReader;
    pending = settingsApi
      .retrieveRow("usersettings", { id: userId, select: USER_SETTINGS_SELECT })
      .then((settings) => {
        winAny[cacheKey] = settings;
        return settings;
      })
      .finally(() => { if (winAny[inflightKey] === pending) delete winAny[inflightKey]; });
    winAny[inflightKey] = pending;
  }
  return pending;
}

async function loadDashboard(dataApi: UxAgentDataApi): Promise<DashboardLoad> {
  const [speakerResult, plannedResult, completedResult, sentResult, settingsResult] = await Promise.allSettled([
    readAllSpeakers(dataApi),
    readAllPlanned(dataApi),
    readRecentCompleted(dataApi),
    readLatestSent(dataApi),
    readUserSettings(dataApi),
  ]);
  const errors: PageErrorView[] = [];
  const speakers = speakerResult.status === "fulfilled" ? speakerResult.value : [];
  const planned = plannedResult.status === "fulfilled" ? plannedResult.value : [];
  const completedAll = completedResult.status === "fulfilled" ? completedResult.value : [];
  const latestSent = sentResult.status === "fulfilled" ? sentResult.value : null;
  const completed = completedAll.slice(0, 5);
  const userSettings = settingsResult.status === "fulfilled" ? settingsResult.value : null;
  if (speakerResult.status === "rejected") errors.push(reportPageError(PAGE_NAME, "loadSpeakers", undefined, speakerResult.reason, "read"));
  if (plannedResult.status === "rejected") errors.push(reportPageError(PAGE_NAME, "loadPlannedHistory", undefined, plannedResult.reason, "read"));
  if (completedResult.status === "rejected") errors.push(reportPageError(PAGE_NAME, "loadCompletedHistory", undefined, completedResult.reason, "read"));
  if (sentResult.status === "rejected") errors.push(reportPageError(PAGE_NAME, "loadLatestSentHistory", undefined, sentResult.reason, "read"));
  if (settingsResult.status === "rejected") errors.push(reportPageError(PAGE_NAME, "loadUserSettings", undefined, settingsResult.reason, "read"));
  return { snapshot: { speakers, planned, completed, latestSent, completedCount: completedAll.length, userSettings }, errors };
}

function currentUserId(): string | undefined {
  try {
    return getXrm()?.Utility?.getGlobalContext?.().userSettings?.userId?.replace(/[{}]/g, "");
  } catch {
    return undefined;
  }
}

function formatNumber(value: number, settings: UserSettingsRow | null, locale: string): string {
  if (!settings?.decimalsymbol || !settings.numberseparator) return value.toLocaleString(locale);
  const sign = value < 0 ? "-" : "";
  const [integer, fraction] = Math.abs(value).toString().split(".");
  const grouped = integer.replace(/\B(?=(\d{3})+(?!\d))/g, settings.numberseparator);
  return sign + grouped + (fraction ? settings.decimalsymbol + fraction : "");
}

function formatDate(value: Date | string | null | undefined, settings: UserSettingsRow | null, locale: string): string {
  if (!value) return "—";
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  const pattern = settings?.dateformatstring;
  const separator = settings?.dateseparator;
  if (!pattern || !separator) return date.toLocaleDateString(locale);
  const year = date.getFullYear();
  const month = date.getMonth() + 1;
  const day = date.getDate();
  return pattern.replace(/[/\-.]/g, separator).replace(/yyyy|yy|MM|M|dd|d/g, (token) => {
    switch (token) {
      case "yyyy": return String(year);
      case "yy": return String(year).slice(-2);
      case "MM": return String(month).padStart(2, "0");
      case "M": return String(month);
      case "dd": return String(day).padStart(2, "0");
      case "d": return String(day);
      default: return token;
    }
  });
}

function formatDateTime(value: Date | string | null | undefined, settings: UserSettingsRow | null, locale: string): string {
  if (!value) return "—";
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return "—";
  return `${formatDate(date, settings, locale)} · ${date.toLocaleTimeString(locale, { hour: "2-digit", minute: "2-digit" })}`;
}

function nextMondayAtNine(): Date {
  const now = new Date();
  const daysUntilMonday = (8 - now.getDay()) % 7 || 7;
  return new Date(now.getFullYear(), now.getMonth(), now.getDate() + daysUntilMonday, 9, 0, 0, 0);
}

function wheelGradient(count: number): string {
  if (count < 1) return "conic-gradient(#E6E2F3 0deg 360deg)";
  const colors = ["#5B3FD3", "#7A68E8", "#9A8BEF", "#B6A9F6", "#D1C9FB", "#7560D9"];
  const segment = 360 / count;
  const stops: string[] = [];
  for (let index = 0; index < count; index += 1) {
    const start = index * segment;
    const end = (index + 1) * segment;
    const color = colors[index % colors.length];
    stops.push(`${color} ${start}deg ${end}deg`);
  }
  return `conic-gradient(${stops.join(", ")})`;
}

type PartialDataWriter = {
  updateRow(
    tableName: "exp_speakers",
    rowId: string,
    row: Partial<WritableTableRow<exp_speakers>>,
  ): Promise<void>;
  updateRow(
    tableName: "exp_icebreakerhistory",
    rowId: string,
    row: Partial<WritableTableRow<exp_icebreakerhistory>>,
  ): Promise<void>;
};

function KpiCard(props: { label: string; value: string; foot: string; icon: React.ReactNode }) {
  const styles = useStyles();
  return (
    <Card className={styles.kpiCard} aria-label={`${props.label}: ${props.value}`}>
      <div className={styles.kpiTop}>
        <Text className={styles.kpiLabel}>{props.label}</Text>
        <div className={styles.kpiIcon} aria-hidden="true">{props.icon}</div>
      </div>
      <Text className={styles.kpiValue}>{props.value}</Text>
      <Text className={styles.kpiFoot}>{props.foot}</Text>
    </Card>
  );
}

function PageBody(props: GeneratedComponentProps) {
  const { dataApi } = props;
  const actionApi = dataApi as unknown as ActionDataApi | undefined;
  const styles = useStyles();
  const language = React.useMemo(() => resolvePageLanguage(), []);
  const locale = language.code;
  const cached = winAny[CACHE_KEY] as DashboardSnapshot | undefined;
  const [data, setData] = React.useState<DashboardState>(() => ({
    snapshot: cached ?? { speakers: [], planned: [], completed: [], latestSent: null, completedCount: 0, userSettings: null },
    errors: [],
    loading: cached === undefined,
  }));
  const [reloadKey, setReloadKey] = React.useState(0);
  const [action, setAction] = React.useState<{ busy: boolean; error: PageErrorView | null; notice: string; operation?: "draw" | "complete" | "reset" }>({
    busy: false,
    error: null,
    notice: "",
  });
  const [navigationError, setNavigationError] = React.useState<PageErrorView | null>(null);
  const [completionForm, setCompletionForm] = React.useState({ activity: "", comments: "" });
  const [refreshCountdown, setRefreshCountdown] = React.useState(30);
  const [drawPending, setDrawPending] = React.useState(false);
  const [prefersReducedMotion, setPrefersReducedMotion] = React.useState(false);
  const [wheelRotation, setWheelRotation] = React.useState(0);
  const [resetBusy, setResetBusy] = React.useState(false);
  const drawBaselineId = React.useRef<string | null>(null);
  const refreshCountdownRef = React.useRef(30);
  const actionPending = React.useRef(false);
  const dataReady = !!dataApi;

  React.useEffect(() => {
    if (!dataReady) return;
    const currentCache = winAny[CACHE_KEY] as DashboardSnapshot | undefined;
    if (currentCache !== undefined) {
      setData((previous) => previous.snapshot === currentCache
        ? { ...previous, loading: false }
        : { snapshot: currentCache, errors: [], loading: false });
      return;
    }
    let cancelled = false;
    let pending = winAny[INFLIGHT_KEY] as Promise<DashboardLoad> | undefined;
    if (!pending) {
      pending = loadDashboard(dataApi)
        .then((result) => {
          if (result.errors.length === 0) winAny[CACHE_KEY] = result.snapshot;
          return result;
        })
        .finally(() => { if (winAny[INFLIGHT_KEY] === pending) delete winAny[INFLIGHT_KEY]; });
      winAny[INFLIGHT_KEY] = pending;
    }
    pending
      .then((result) => {
        if (!cancelled) setData({ ...result, loading: false });
      })
      .catch((failure) => {
        if (!cancelled) {
          setData({
            snapshot: { speakers: [], planned: [], completed: [], latestSent: null, completedCount: 0, userSettings: null },
            errors: [reportPageError(PAGE_NAME, "loadDashboard", undefined, failure, "read")],
            loading: false,
          });
        }
      });
    return () => { cancelled = true; };
    // The host supplies a new dataApi reference on each render; readiness is the only dependency.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dataReady, reloadKey]);

  React.useEffect(() => {
    if (!dataReady) return;
    const timerId = window.setInterval(() => {
      const next = refreshCountdownRef.current - 1;
      if (next > 0) {
        refreshCountdownRef.current = next;
        setRefreshCountdown(next);
        return;
      }
      refreshCountdownRef.current = 30;
      setRefreshCountdown(30);
      if (actionPending.current || winAny[INFLIGHT_KEY]) return;
      invalidateAndReload();
    }, 1000);
    return () => window.clearInterval(timerId);
    // The timer reads mutable refs so it can avoid overlapping reads and writes.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [dataReady]);

  React.useEffect(() => {
    const latestSentId = data.snapshot.latestSent?.exp_icebreakerhistoryid ?? null;
    if (drawPending && latestSentId !== null && latestSentId !== drawBaselineId.current) {
      setDrawPending(false);
    }
  }, [drawPending, data.snapshot.latestSent?.exp_icebreakerhistoryid]);

  React.useEffect(() => {
    if (typeof window.matchMedia !== "function") return;
    const mediaQuery = window.matchMedia("(prefers-reduced-motion: reduce)");
    const updatePreference = () => setPrefersReducedMotion(mediaQuery.matches);
    updatePreference();
    if (typeof mediaQuery.addEventListener === "function") {
      mediaQuery.addEventListener("change", updatePreference);
      return () => mediaQuery.removeEventListener("change", updatePreference);
    }
    mediaQuery.addListener(updatePreference);
    return () => mediaQuery.removeListener(updatePreference);
  }, []);

  React.useEffect(() => {
    const latestSentId = data.snapshot.latestSent?.exp_icebreakerhistoryid ?? null;
    if (!drawPending || prefersReducedMotion || (latestSentId !== null && latestSentId !== drawBaselineId.current)) return;
    const advanceWheel = () => setWheelRotation((rotation) => rotation + 360);
    advanceWheel();
    const timerId = window.setInterval(advanceWheel, 900);
    return () => window.clearInterval(timerId);
  }, [drawPending, prefersReducedMotion, data.snapshot.latestSent?.exp_icebreakerhistoryid]);

  const activeSpeakers = data.snapshot.speakers;
  const eligibleSpeakers = activeSpeakers.filter((speaker) => speaker.exp_statut === 791170001);
  const activeCount = activeSpeakers.length;
  const plannedRows = data.snapshot.planned;
  const completedRows = data.snapshot.completed;
  const closeTarget = plannedRows.reduce<HistoryRow | null>((closest, row) => {
    if (!closest) return row;
    const currentDate = row.exp_date instanceof Date ? row.exp_date.getTime() : new Date(row.exp_date).getTime();
    const closestDate = closest.exp_date instanceof Date ? closest.exp_date.getTime() : new Date(closest.exp_date).getTime();
    if (!Number.isFinite(currentDate)) return closest;
    if (!Number.isFinite(closestDate) || Math.abs(currentDate - Date.now()) < Math.abs(closestDate - Date.now())) return row;
    return closest;
  }, null);
  const closeLookup = closeTarget
    ? normalizeLookupReference(closeTarget, "_exp_speaker_value", ["exp_speakers"])
    : { state: "unset" as const };

  React.useEffect(() => {
    setCompletionForm({ activity: "", comments: "" });
  }, [closeTarget?.exp_icebreakerhistoryid]);

  function invalidateAndReload(): void {
    delete winAny[CACHE_KEY];
    delete winAny[INFLIGHT_KEY];
    refreshCountdownRef.current = 30;
    setRefreshCountdown(30);
    setData((previous) => ({ ...previous, loading: true, errors: [] }));
    setReloadKey((key) => key + 1);
  }

  function retryLoad(): void {
    delete winAny[CACHE_KEY];
    delete winAny[INFLIGHT_KEY];
    refreshCountdownRef.current = 30;
    setRefreshCountdown(30);
    setData((previous) => ({ ...previous, loading: true, errors: [] }));
    setAction({ busy: false, error: null, notice: "" });
    setReloadKey((key) => key + 1);
  }

  async function openEntityList(entityName: string): Promise<void> {
    setNavigationError(null);
    const navigation = getXrm()?.Navigation;
    if (!navigation) {
      setNavigationError(reportPageError(PAGE_NAME, "openEntityList", undefined, new Error("Navigation is unavailable."), "read"));
      return;
    }
    try {
      await navigation.navigateTo({ pageType: "entitylist", entityName });
    } catch (error) {
      setNavigationError(reportPageError(PAGE_NAME, "openEntityList", undefined, error, "read"));
    }
  }

  async function drawSpeaker(isManualRetry = false): Promise<void> {
    if (actionPending.current || data.errors.some((error) => error.operation !== "loadUserSettings") || (!isManualRetry && action.error) || eligibleSpeakers.length === 0) return;
    if (typeof actionApi?.executeAction !== "function") {
      setDrawPending(false);
      setAction({ busy: false, error: reportPageError(PAGE_NAME, "drawIceBreaker", undefined, new Error("Custom API actions are unavailable in this Power Apps runtime."), "write"), notice: "" });
      return;
    }
    if (!isManualRetry) {
      drawBaselineId.current = data.snapshot.latestSent?.exp_icebreakerhistoryid ?? null;
      setDrawPending(true);
    }
    actionPending.current = true;
    setAction({ busy: true, error: null, notice: "", operation: "draw" });
    try {
      const result = await actionApi.executeAction({ name: "exp_TriggerIceBreakerSelection" });
      if (result.indeterminate) {
        setAction({ busy: false, error: reportPageError(PAGE_NAME, "drawIceBreaker", result.error?.code, result.error?.message, "indeterminate"), notice: "" });
        return;
      }
      if (result.error?.code === "indeterminate") {
        setAction({ busy: false, error: reportPageError(PAGE_NAME, "drawIceBreaker", result.error.code, result.error.message, "indeterminate"), notice: "" });
        return;
      }
      if (!result.ok && result.error?.code === "network") {
        setAction({
          busy: false,
          error: reportPageError(PAGE_NAME, "drawIceBreaker", result.error.code, result.error.message, "network", () => { void drawSpeaker(true); }),
          notice: "",
        });
        return;
      }
      if (!result.ok) {
        setDrawPending(false);
        setAction({ busy: false, error: reportPageError(PAGE_NAME, "drawIceBreaker", result.error?.code, result.error?.message, "write"), notice: "" });
        return;
      }
      setAction({ busy: false, error: null, notice: translate("drawSuccess"), operation: "draw" });
      invalidateAndReload();
    } catch (error) {
      setDrawPending(false);
      setAction({ busy: false, error: reportPageError(PAGE_NAME, "drawIceBreaker", undefined, error, "write"), notice: "" });
    } finally {
      actionPending.current = false;
    }
  }

  async function completeIceBreaker(): Promise<void> {
    if (actionPending.current || data.errors.some((error) => error.operation !== "loadUserSettings") || action.error || !closeTarget || closeLookup.state !== "resolved") return;
    actionPending.current = true;
    setAction({ busy: true, error: null, notice: "", operation: "complete" });
    try {
      const patchApi = dataApi as unknown as PartialDataWriter;
      await patchApi.updateRow("exp_icebreakerhistory", closeTarget.exp_icebreakerhistoryid, {
        exp_activite: completionForm.activity.trim(),
        exp_commentaires: completionForm.comments.trim(),
        exp_statut: 791170002,
      });
      await patchApi.updateRow("exp_speakers", closeLookup.reference.id, {
        exp_statut: 791170003,
        exp_derniererealisation: new Date(),
      });
      setAction({ busy: false, error: null, notice: translate("closeSuccess"), operation: "complete" });
      setCompletionForm({ activity: "", comments: "" });
      invalidateAndReload();
    } catch (error) {
      setAction({ busy: false, error: reportPageError(PAGE_NAME, "completeIceBreaker", undefined, error, "write"), notice: "", operation: "complete" });
    } finally {
      actionPending.current = false;
    }
  }

  async function resetActiveSpeakers(): Promise<void> {
    if (actionPending.current || data.errors.some((error) => error.operation !== "loadUserSettings")) return;
    const speakersToReset = activeSpeakers.filter((speaker) => speaker.exp_statut !== 791170001);
    if (speakersToReset.length === 0) {
      setAction({ busy: false, error: null, notice: translate("resetNoop"), operation: "reset" });
      return;
    }
    actionPending.current = true;
    setResetBusy(true);
    setAction({ busy: true, error: null, notice: "", operation: "reset" });
    try {
      const patchApi = dataApi as unknown as PartialDataWriter;
      const results = await Promise.allSettled(speakersToReset.map((speaker) =>
        patchApi.updateRow("exp_speakers", speaker.exp_speakersid, { exp_statut: 791170001 }),
      ));
      const failureCount = results.filter((result) => result.status === "rejected").length;
      if (failureCount > 0) {
        throw new Error(`Reset failed for ${failureCount} of ${speakersToReset.length} active speakers.`);
      }
      setAction({ busy: false, error: null, notice: translate("resetSuccess"), operation: "reset" });
      invalidateAndReload();
    } catch (error) {
      setAction({ busy: false, error: reportPageError(PAGE_NAME, "resetActiveSpeakers", undefined, error, "write"), notice: "", operation: "reset" });
      invalidateAndReload();
    } finally {
      actionPending.current = false;
      setResetBusy(false);
    }
  }

  const settings = data.snapshot.userSettings;
  const wheelBackground = wheelGradient(eligibleSpeakers.length);
  const scheduledFor = nextMondayAtNine();
  const latestSent = data.snapshot.latestSent;
  const latestSentId = latestSent?.exp_icebreakerhistoryid ?? null;
  const awaitingDrawResult = drawPending && (latestSentId === null || latestSentId === drawBaselineId.current);
  const wheelCenterHeading = awaitingDrawResult
    ? translate("wheelPendingLabel")
    : latestSent
      ? translate("wheelResultLabel")
      : "";
  const wheelCenterText = awaitingDrawResult
    ? translate("wheelPendingValue")
    : latestSent
      ? getHistorySpeakerLabel(latestSent)
      : translate("eligible");
  const wheelAccessibleLabel = awaitingDrawResult
    ? translate("drawPendingStatus")
    : latestSent
      ? translate("latestSentSpeaker").replace("{speaker}", getHistorySpeakerLabel(latestSent))
      : `${translate("eligible")}: ${eligibleSpeakers.map((speaker) => speaker.exp_name).join(", ")}`;
  const hasInitialRows = activeSpeakers.length > 0 || plannedRows.length > 0 || completedRows.length > 0 || latestSent !== null;
  const noDataYet = data.loading && !hasInitialRows;

  return (
    <div style={{ ...themeToVars(BRAND_THEME), height: "100%", overflowY: "auto", boxSizing: "border-box", minHeight: 0, position: "relative", contain: "layout" }}>
      <div className={styles.root} dir={language.isRtl ? "rtl" : "ltr"}>
        <aside className={styles.sidebar}>
          <div className={styles.brand}>
            <div className={styles.brandMark} aria-hidden="true">e</div>
            <div className={styles.brandText}>
              <Text className={styles.brandName}>{translate("appName")}</Text>
            </div>
          </div>
          <nav className={styles.nav} aria-label={translate("navLabel")}>
            <Button appearance="transparent" className={styles.navLink} aria-current="page" icon={<HomeRegular />}>
              {translate("navHome")}
            </Button>
            <Button appearance="transparent" className={styles.navLink} icon={<PeopleRegular />} onClick={() => void openEntityList("exp_speakers")}>
              {translate("navSpeakers")}
            </Button>
            <Button appearance="transparent" className={styles.navLink} icon={<HistoryRegular />} onClick={() => void openEntityList("exp_icebreakerhistory")}>
              {translate("navHistory")}
            </Button>
          </nav>
          <Text className={styles.sidebarFoot}>{translate("sidebarCaption")}</Text>
        </aside>

        <main className={styles.main}>
          <header className={styles.header}>
            <div className={styles.refreshControls}>
              <Button
                appearance="subtle"
                className={styles.refreshButton}
                icon={<ArrowSyncRegular />}
                aria-label={translate("refreshLabel")}
                disabled={action.busy || data.loading}
                onClick={retryLoad}
              >
                {translate("refresh")}
              </Button>
              <Text className={styles.refreshCountdown}>
                {translate("autoRefreshCountdown").replace("{seconds}", formatNumber(refreshCountdown, settings, locale))}
              </Text>
            </div>
          </header>

          <div className={styles.errorStack}>
            {data.errors.map((error) => <PageErrorBar key={error.operation} error={{ ...error, retry: retryLoad }} />)}
            <PageErrorBar error={action.error} />
            <PageErrorBar error={navigationError} />
            {settings === null && data.errors.some((error) => error.operation === "loadUserSettings") && (
              <MessageBar intent="warning"><MessageBarBody>{translate("settingsLoadFailed")}</MessageBarBody></MessageBar>
            )}
          </div>

          {noDataYet ? (
            <div className={styles.emptyState}><Spinner label={translate("loading")} /></div>
          ) : (
            <>

              {data.loading && (
                <div className={styles.inlineLoading}><Spinner size="tiny" label={translate("loading")} /></div>
              )}

              <section className={styles.contentGrid} aria-label={translate("selectionReady")}>
                <Card className={styles.panel}>
                  <div className={styles.selectionLayout}>
                    <div className={styles.wheelColumn}>
                      {eligibleSpeakers.length > 0 || latestSent !== null ? (
                        <>
                          <div className={styles.wheelWrap} role="img" aria-label={wheelAccessibleLabel}>
                            <div className={styles.wheelPointer} aria-hidden="true" />
                            <div
                              className={styles.wheel}
                              style={{
                                background: wheelBackground,
                                transform: `rotate(${wheelRotation}deg)`,
                                transition: drawPending && !prefersReducedMotion ? "transform 900ms linear" : "none",
                              }}
                            />
                            <div className={styles.wheelCenter}>
                              {wheelCenterHeading && <Text className={styles.wheelCenterLabel}>{wheelCenterHeading}</Text>}
                              <Text className={styles.wheelCenterName}>{wheelCenterText}</Text>
                            </div>
                          </div>
                          <Text className={styles.eligibleCaption}>{formatNumber(eligibleSpeakers.length, settings, locale)} · {translate("eligible")}</Text>
                          <div className={styles.eligibleNames} aria-label={translate("eligible")}>
                            {eligibleSpeakers.slice(0, 8).map((speaker) => (
                              <Text key={speaker.exp_speakersid} className={styles.eligibleChip}>
                                {speaker.exp_name?.trim() || translate("noSpeakerName")}
                              </Text>
                            ))}
                            {eligibleSpeakers.length > 8 && (
                              <Text className={styles.eligibleChip}>+{formatNumber(eligibleSpeakers.length - 8, settings, locale)}</Text>
                            )}
                          </div>
                        </>
                      ) : (
                        <Text className={styles.emptyState}>{translate("wheelEmpty")}</Text>
                      )}
                    </div>
                    <div className={styles.actionColumn}>
                      <Text className={styles.actionEyebrow}>{translate("eligible")}</Text>
                      <Text className={styles.actionTitle}>{translate("selectionReady")}</Text>
                      <Text className={styles.muted}>{translate("drawSchedule")}</Text>
                      <Button
                        appearance="primary"
                        className={styles.primaryAction}
                        icon={action.busy && action.operation === "draw" ? <Spinner size="tiny" /> : <PlayRegular />}
                        disabled={drawPending || action.busy || action.error !== null || data.loading || data.errors.some((error) => error.operation !== "loadUserSettings") || eligibleSpeakers.length === 0}
                        onClick={() => void drawSpeaker()}
                      >
                        {action.busy && action.operation === "draw" ? translate("spinning") : translate("spinAction")}
                      </Button>
                      {eligibleSpeakers.length === 0 && <Text className={styles.muted}>{translate("drawUnavailable")}</Text>}
                      {action.notice && <Text className={styles.resultLine}><CheckmarkCircleRegular />{action.notice}</Text>}
                      <div className={styles.drawStatus} role="status" aria-live="polite">
                        {drawPending && (latestSent?.exp_icebreakerhistoryid ?? null) === drawBaselineId.current
                          ? translate("drawPendingStatus")
                          : latestSent
                            ? translate("latestSentSpeaker").replace("{speaker}", getHistorySpeakerLabel(latestSent))
                            : translate("noSentSpeaker")}
                      </div>
                      <Button
                        appearance="outline"
                        className={styles.secondaryAction}
                        icon={resetBusy ? <Spinner size="tiny" /> : <ArrowSyncRegular />}
                        disabled={action.busy || action.error !== null || data.loading || data.errors.some((error) => error.operation !== "loadUserSettings")}
                        onClick={() => void resetActiveSpeakers()}
                      >
                        {resetBusy ? translate("resettingSpeakers") : translate("resetSpeakers")}
                      </Button>
                    </div>
                  </div>
                </Card>

                <Card className={styles.panel}>
                  <div className={styles.panelHeader}>
                    <Text as="h2" className={styles.sectionTitle}>{translate("closeTitle")}</Text>
                    <Text className={styles.sectionDescription}>{translate("closeDescription")}</Text>
                  </div>
                  {closeTarget ? (
                    <>
                      <div>
                        <Text className={styles.statusName}>{getHistorySpeakerLabel(closeTarget)}</Text>
                        <Text className={styles.statusDate}>{formatDateTime(closeTarget.exp_date, settings, locale)}</Text>
                      </div>
                      <div className={styles.completionFields}>
                        <div className={styles.fieldGroup}>
                          <Text className={styles.fieldLabel}>{translate("activityLabel")}</Text>
                          <Input
                            value={completionForm.activity}
                            aria-label={translate("activityLabel")}
                            placeholder={translate("activityPlaceholder")}
                            onChange={(_, value) => setCompletionForm((previous) => ({ ...previous, activity: value.value }))}
                          />
                        </div>
                        <div className={styles.fieldGroup}>
                          <Text className={styles.fieldLabel}>{translate("commentLabel")}</Text>
                          <Textarea
                            value={completionForm.comments}
                            aria-label={translate("commentLabel")}
                            placeholder={translate("commentPlaceholder")}
                            resize="vertical"
                            onChange={(_, value) => setCompletionForm((previous) => ({ ...previous, comments: value.value }))}
                          />
                        </div>
                      </div>
                      {closeLookup.state !== "resolved" && (
                        <MessageBar intent="warning"><MessageBarBody>{translate("unresolvedSpeaker")}</MessageBarBody></MessageBar>
                      )}
                      <Button
                        appearance="primary"
                        className={styles.completionButton}
                        icon={action.busy && action.operation === "complete" ? <Spinner size="tiny" /> : <CheckmarkCircleRegular />}
                        disabled={action.busy || action.error !== null || data.loading || data.errors.some((error) => error.operation !== "loadUserSettings") || closeLookup.state !== "resolved"}
                        onClick={() => void completeIceBreaker()}
                      >
                        {action.busy && action.operation === "complete" ? translate("saving") : translate("markDone")}
                      </Button>
                    </>
                  ) : (
                    <Text className={styles.emptyState}>{translate("noPlanned")}</Text>
                  )}
                </Card>
              </section>

              <section className={styles.tablePanel} aria-labelledby="recent-heading">
                <div className={styles.panelHeader}>
                  <Text as="h2" id="recent-heading" className={styles.sectionTitle}>{translate("recentTitle")}</Text>
                </div>
                {completedRows.length > 0 ? (
                  <div className={styles.tableWrap}>
                    <table className={styles.table}>
                      <thead>
                        <tr>
                          <th className={styles.tableHeadCell}>{translate("speakerColumn")}</th>
                          <th className={styles.tableHeadCell}>{translate("dateColumn")}</th>
                          <th className={styles.tableHeadCell}>{translate("activityColumn")}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {completedRows.map((row) => {
                          const lookup = normalizeLookupReference(row, "_exp_speaker_value", ["exp_speakers"]);
                          const speakerLabel = lookup.state === "resolved" ? getHistorySpeakerLabel(row) : translate("speakerUnavailable");
                          return (
                            <tr key={row.exp_icebreakerhistoryid}>
                              <td className={`${styles.tableCell} ${styles.tableStrong}`}>{speakerLabel}</td>
                              <td className={styles.tableCell}>{formatDate(row.exp_date, settings, locale)}</td>
                              <td className={styles.tableCell}>{row.exp_activite?.trim() || "—"}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <Text className={styles.emptyState}>{translate("noRecent")}</Text>
                )}
              </section>

              {plannedRows.length > 1 && (
                <section className={`${styles.panel} ${styles.alertPanel}`} aria-labelledby="planned-alert-heading">
                  <div className={styles.panelHeader}>
                    <Text as="h2" id="planned-alert-heading" className={styles.alertHeading}>
                      <WarningRegular aria-hidden="true" />{translate("plannedAlertTitle")}
                    </Text>
                    <Text className={styles.alertCopy}>{translate("plannedAlertDescription")}</Text>
                    <Text className={styles.alertCount}>{translate("plannedCount").replace("{count}", formatNumber(plannedRows.length, settings, locale))}</Text>
                  </div>
                  <div>
                    {plannedRows.map((row) => (
                      <div key={row.exp_icebreakerhistoryid} className={styles.statusRow}>
                        <Text className={styles.statusName}>{getHistorySpeakerLabel(row)}</Text>
                        <Text className={styles.statusDate}>{formatDateTime(row.exp_date, settings, locale)}</Text>
                      </div>
                    ))}
                  </div>
                  <Button
                    appearance="outline"
                    className={styles.alertButton}
                    icon={<ArrowRightRegular />}
                    onClick={() => void openEntityList("exp_icebreakerhistory")}
                  >
                    {translate("openPlans")}
                  </Button>
                </section>
              )}

              <Text className={styles.muted} aria-label={translate("drawSchedule")}>
                <CalendarLtrRegular /> {translate("drawSchedule")} · {formatDate(scheduledFor, settings, locale)}
              </Text>
            </>
          )}
        </main>
      </div>
    </div>
  );
}

const GeneratedComponent = (props: GeneratedComponentProps) => (
  <PageErrorBoundary>
    <PageBody {...props} />
  </PageErrorBoundary>
);

export default GeneratedComponent;
