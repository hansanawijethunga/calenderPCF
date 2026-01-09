
import * as React from "react";
import type { FC, MouseEvent } from "react";
import {
  FluentProvider,
  webLightTheme,
  Toolbar,
  ToolbarButton,
  TabList,
  Tab,
  Divider,
  Text,
  Tooltip,
  Badge,
  tokens,
  makeStyles,
  mergeClasses,
} from "@fluentui/react-components";

/* ----------------------------- PCF dataset types ---------------------------- */
interface ColumnDef {
  name: string;
  displayName: string;
}
interface EntityRecord {
  getValue: (columnName: string) => unknown;
}
interface DataSetShape {
  columns: ColumnDef[];
  sortedRecordIds: string[];
  records: Record<string, EntityRecord>;
}

/* ------------------------------- Component API ------------------------------ */
export interface CalendarProps {
  dataset: DataSetShape;
  fromColumn: string; // start datetime column name
  toColumn: string; // end datetime column name
  titleColumn: string;
  subtitleColumn: string;
  typeColumn: string;
  onSelect?: (selectedRecordIds: string[]) => void;
  containerHeight?: number;
}

/* ------------------------------ Internal models ----------------------------- */
type CalendarView = "month" | "week" | "day";

interface CalendarEvent {
  id: string;
  title: string;
  subtitle?: string;
  type?: string;
  start: Date;
  end: Date;
}

/* ----------------------------- Date/Time utilities -------------------------- */
const DAY_MS = 24 * 60 * 60 * 1000;
const WORK_START_MIN = 9 * 60; // 09:00
const WORK_END_MIN = 18 * 60; // 18:00
const WORK_TOTAL_MIN = WORK_END_MIN - WORK_START_MIN; // 540
const HEADER_HEIGHT = 48; // Height of day column headers in pixels

const toDate = (value: unknown): Date | null => {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === "string") {
    const d = new Date(value);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
};

const isSameDay = (a: Date, b: Date): boolean =>
  a.getFullYear() === b.getFullYear() &&
  a.getMonth() === b.getMonth() &&
  a.getDate() === b.getDate();

const startOfDay = (d: Date): Date =>
  new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);

const addDays = (d: Date, delta: number): Date =>
  new Date(d.getFullYear(), d.getMonth(), d.getDate() + delta, d.getHours(), d.getMinutes(), d.getSeconds(), d.getMilliseconds());

const clamp = (val: number, min: number, max: number): number =>
  Math.max(min, Math.min(max, val));

const minutesFromNine = (d: Date): number => d.getHours() * 60 + d.getMinutes() - WORK_START_MIN;

const formatHourLabel = (hour24: number): string => {
  const h = ((hour24 + 11) % 12) + 1;
  const ampm = hour24 >= 12 ? "pm" : "am";
  return `${h}:00 ${ampm}`;
};

const formatShortHourLabel = (hour24: number): string => {
  const h = ((hour24 + 11) % 12) + 1;
  const ampm = hour24 >= 12 ? "pm" : "am";
  return `${h}${ampm}`;
};

const pad2 = (n: number): string => (n < 10 ? `0${n}` : `${n}`);

const formatTimeRange = (start: Date, end: Date): string => {
  const sH = ((start.getHours() + 11) % 12) + 1;
  const sM = pad2(start.getMinutes());
  const sAm = start.getHours() >= 12 ? "pm" : "am";
  const eH = ((end.getHours() + 11) % 12) + 1;
  const eM = pad2(end.getMinutes());
  const eAm = end.getHours() >= 12 ? "pm" : "am";
  return `${sH}:${sM}${sAm}–${eH}:${eM}${eAm}`;
};

/* -------------------------- Event normalization helpers --------------------- */
const getString = (value: unknown): string | undefined => {
  if (value === null || value === undefined) return undefined;
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (typeof value === "object") {
    const anyVal = value as { name?: string; label?: string; toString?: () => string };
    if (anyVal.name) return anyVal.name;
    if (anyVal.label) return anyVal.label;
    if (anyVal.toString) return anyVal.toString();
  }
  return undefined;
};

const normalizeEvents = (
  dataset: DataSetShape,
  titleColumn: string,
  subtitleColumn: string,
  typeColumn: string,
  fromColumn: string,
  toColumn: string
): CalendarEvent[] => {
  const ids = dataset?.sortedRecordIds ?? [];
  const events: CalendarEvent[] = [];
  for (const id of ids) {
    const record = dataset.records[id];
    if (!record) continue;
    const start = toDate(record.getValue(fromColumn));
    const end = toDate(record.getValue(toColumn));
    if (!start || !end || isNaN(start.getTime()) || isNaN(end.getTime())) continue;
    const title = getString(record.getValue(titleColumn)) ?? "(Untitled)";
    const subtitle = getString(record.getValue(subtitleColumn));
    const type = getString(record.getValue(typeColumn));
    events.push({ id, title, subtitle, type, start, end });
  }
  return events;
};

/* ------------------------------- Lanes algorithm ---------------------------- */
interface PositionedEvent extends CalendarEvent {
  // Segment adjusted/clamped to the current day
  displayStart: Date;
  displayEnd: Date;
  topPct: number;
  heightPct: number;
  lane: number;
  laneCount: number;
}

const computeDaySegments = (
  day: Date,
  events: CalendarEvent[]
): (CalendarEvent & { displayStart: Date; displayEnd: Date })[] => {
  const dayStart = startOfDay(day);
  const dayEnd = addDays(dayStart, 1);
  return events
    .map((e) => {
      const s = e.start < dayStart ? dayStart : e.start;
      const ed = e.end > dayEnd ? dayEnd : e.end;
      if (ed <= s) return null;
      return { ...e, displayStart: s, displayEnd: ed };
    })
    .filter((x): x is CalendarEvent & { displayStart: Date; displayEnd: Date } => x !== null);
};

interface Lane { end: Date }

const computeLanes = (segments: (CalendarEvent & { displayStart: Date; displayEnd: Date })[]): (PositionedEvent)[] => {
  const items = [...segments].sort((a, b) => a.displayStart.getTime() - b.displayStart.getTime());
  const lanes: Lane[] = [];

  const placed: PositionedEvent[] = [];
  for (const it of items) {
    let laneIndex = -1;
    for (let i = 0; i < lanes.length; i++) {
      if (lanes[i].end <= it.displayStart) {
        laneIndex = i;
        break;
      }
    }
    if (laneIndex === -1) {
      lanes.push({ end: it.displayEnd });
      laneIndex = lanes.length - 1;
    } else {
      lanes[laneIndex].end = it.displayEnd;
    }
    
    // Calculate position as percentage of work hours (9am-6pm)
    const startMin = minutesFromNine(it.displayStart);
    const endMin = minutesFromNine(it.displayEnd);
    
    // Clamp to visible work hours
    const clampedStartMin = clamp(startMin, 0, WORK_TOTAL_MIN);
    const clampedEndMin = clamp(endMin, 0, WORK_TOTAL_MIN);
    
    // Calculate actual duration in minutes within work hours
    const durMin = Math.max(15, clampedEndMin - clampedStartMin); // minimum 15 min visual
    
    const topPct = (clampedStartMin / WORK_TOTAL_MIN) * 100;
    const heightPct = (durMin / WORK_TOTAL_MIN) * 100;

    placed.push({
      ...it,
      displayStart: it.displayStart,
      displayEnd: it.displayEnd,
      topPct,
      heightPct,
      lane: laneIndex,
      laneCount: lanes.length,
    });
  }
  // Now set final laneCount for each (same for all in that compute run)
  const finalCount = lanes.length;
  return placed.map((p) => ({ ...p, laneCount: finalCount }));
};

/* ------------------------------ Color mapping ------------------------------- */
const getTypeColors = (type?: string): { bg: string; fg: string; border?: string } => {
  // Use Fluent tokens for consistency across themes
  if (!type) {
    return {
      bg: tokens.colorNeutralBackground3,
      fg: tokens.colorNeutralForeground1,
      border: tokens.colorNeutralStroke2,
    };
  }
  const t = type.toLowerCase();
  if (t.includes("appointment")) {
    return {
      bg: tokens.colorBrandBackground2,
      fg: tokens.colorNeutralForegroundInverted,
      border: tokens.colorBrandStroke1,
    };
  }
  if (t.includes("support")) {
    return {
      bg: tokens.colorStatusWarningBackground1,
      fg: tokens.colorNeutralForeground1,
      border: tokens.colorStatusWarningForeground1,
    };
  }
  if (t.includes("training") || t.includes("workshop")) {
    return {
      bg: tokens.colorPaletteGreenBackground2,
      fg: tokens.colorNeutralForegroundInverted,
      border: tokens.colorPaletteGreenBorderActive,
    };
  }
  return {
    bg: tokens.colorNeutralBackground3,
    fg: tokens.colorNeutralForeground1,
    border: tokens.colorNeutralStroke2,
  };
};

/* ---------------------------------- Styles ---------------------------------- */
const useStyles = makeStyles({
  root: {
    height: "100%",
    width: "100%",
    display: "flex",
    flexDirection: "column",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  header: {
    display: "grid",
    gridTemplateColumns: "1fr auto 1fr",
    alignItems: "center",
    gap: "8px",
    padding: "8px 12px",
  },
  headerLeft: { display: "flex", alignItems: "center", gap: "8px" },
  headerCenter: { display: "flex", justifyContent: "center", alignItems: "center" },
  headerRight: { display: "flex", justifyContent: "flex-end", alignItems: "center" },
  periodTitle: { fontWeight: 600 },

  content: {
    flex: 1,
    minHeight: "200px",
    display: "flex",
    flexDirection: "column",
  },

  /* Month grid */
  monthGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(7, minmax(0, 1fr))",
    gridAutoRows: "minmax(120px, 1fr)",
    borderTop: `1px solid ${tokens.colorNeutralStroke2}`,
    borderLeft: `1px solid ${tokens.colorNeutralStroke2}`,
  },
  dayCell: {
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    padding: "8px",
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    backgroundColor: tokens.colorNeutralBackground1,
  },
  dayHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  outsideMonth: { color: tokens.colorNeutralForeground3 },

  chip: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    borderRadius: tokens.borderRadiusMedium,
    padding: "4px 8px",
    boxShadow: tokens.shadow4,
    cursor: "pointer",
    overflow: "hidden",
  },
  chipTitle: {
    fontSize: "12px",
    fontWeight: 600,
    color: tokens.colorNeutralForeground1,
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  chipTime: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    whiteSpace: "nowrap",
  },
  chipMore: {
    fontSize: "12px",
    color: tokens.colorNeutralForeground3,
    marginTop: "2px",
  },

  /* Week/Day shared */
  schedulerRoot: {
    display: "grid",
    gridTemplateColumns: "64px 1fr",
    height: "100%",
    overflow: "hidden",
  },
  leftRail: {
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground2,
    position: "relative",
    paddingTop: `${HEADER_HEIGHT}px`, // Add padding to align with day headers
  },
  leftRailHour: {
    height: "calc(100% / 9)", // 9 visible hours
    display: "flex",
    alignItems: "flex-start",
    padding: "2px 6px",
    color: tokens.colorNeutralForeground3,
    fontSize: "12px",
  },
  gridArea: {
    position: "relative",
    overflow: "auto",
    backgroundColor: tokens.colorNeutralBackground1,
  },

  /* Week columns */
  weekColumns: {
    display: "grid",
    gridTemplateColumns: "repeat(7, minmax(140px, 1fr))",
    height: "100%",
    position: "relative",
  },
  dayColumn: {
    position: "relative",
    borderRight: `1px solid ${tokens.colorNeutralStroke2}`,
    backgroundColor: tokens.colorNeutralBackground1,
    // Prevent overflow into adjacent columns
    overflow: "hidden",
  },
  dayColumnHeader: {
    position: "sticky",
    top: 0,
    zIndex: 3,
    backgroundColor: tokens.colorNeutralBackground1,
    borderBottom: `1px solid ${tokens.colorNeutralStroke2}`,
    padding: "12px 8px",
    display: "flex",
    alignItems: "center",
    gap: "8px",
    height: `${HEADER_HEIGHT}px`,
    boxSizing: "border-box",
  },
  /* Hour & quarter lines overlay for Week */
  overlayLines: {
    position: "absolute",
    left: 0,
    right: 0,
    top: `${HEADER_HEIGHT}px`, // Start below headers
    bottom: 0,
    pointerEvents: "none",
    zIndex: 1,
  },
  hourLine: {
    position: "absolute",
    left: 0,
    right: 0,
    height: "1px",
    backgroundColor: tokens.colorNeutralStroke2,
  },
  quarterLine: {
    position: "absolute",
    left: 0,
    right: 0,
    height: "1px",
    backgroundColor: tokens.colorNeutralStroke2,
    opacity: 0.35,
  },

  /* Events container - positioned below header */
  eventsContainer: {
    position: "absolute",
    top: `${HEADER_HEIGHT}px`,
    left: 0,
    right: 0,
    bottom: 0,
    // Prevent events from overflowing
    overflow: "hidden",
  },

  /* Events */
  eventCard: {
    position: "absolute",
    borderRadius: tokens.borderRadiusMedium,
    boxShadow: tokens.shadow4,
    padding: "6px 8px",
    display: "flex",
    flexDirection: "column",
    gap: "4px",
    cursor: "pointer",
    overflow: "hidden",
    zIndex: 2,
    // Add small margin to prevent touching borders
    boxSizing: "border-box",
  },
  eventTitle: {
    fontSize: "12px",
    fontWeight: 600,
    overflow: "hidden",
    whiteSpace: "nowrap",
    textOverflow: "ellipsis",
  },
  eventSubtitle: {
    fontSize: "11px",
    color: tokens.colorNeutralForeground3,
    overflow: "hidden",
    whiteSpace: "nowrap",
    textOverflow: "ellipsis",
  },

  /* Day view specific gridlines (stronger hour + sublines) */
  dayGridlines: {
    position: "absolute",
    left: 0,
    right: 0,
    top: `${HEADER_HEIGHT}px`, // Start below header for day view too
    bottom: 0,
    pointerEvents: "none",
    zIndex: 1,
  },

  /* Day view single column wrapper */
  dayColumnWrapper: {
    position: "relative",
    height: "100%",
  },

  /* Responsive rules */
  "@media (max-width: 1199px)": {
    weekColumns: {
      gridTemplateColumns: "repeat(7, minmax(120px, 1fr))",
    },
    leftRail: {
      width: "56px",
    },
  },
  "@media (max-width: 899px)": {
    weekColumns: {
      gridTemplateColumns: "repeat(7, minmax(100px, 1fr))",
    },
    leftRail: {
      width: "48px",
    },
    leftRailHour: {
      fontSize: "11px",
    },
  },
  "@media (max-width: 599px)": {
    schedulerRoot: {
      gridTemplateColumns: "0px 1fr", // hide left rail in ultra compact
    },
    leftRail: { display: "none" },
    dayColumnHeader: { padding: "6px" },
  },
});

/* ------------------------------- Month helpers ------------------------------ */
const getFirstMondayOfMatrix = (focus: Date): Date => {
  const first = new Date(focus.getFullYear(), focus.getMonth(), 1);
  const dow = first.getDay(); // 0 Sun .. 6 Sat
  // We want Monday as start; compute delta from Monday
  const deltaToMonday = ((dow + 6) % 7); // converts Sun(0)->6, Mon(1)->0, ...
  return addDays(first, -deltaToMonday);
};

const buildMonthMatrix = (focus: Date): Date[] => {
  const start = getFirstMondayOfMatrix(focus);
  const days: Date[] = [];
  for (let i = 0; i < 42; i++) {
    days.push(addDays(start, i));
  }
  return days;
};

const getMonthTitle = (date: Date): string =>
  `${date.toLocaleString(undefined, { month: "long" })} ${date.getFullYear()}`;

const getWeekTitle = (monday: Date): string => {
  const sunday = addDays(monday, 6);
  const opts: Intl.DateTimeFormatOptions = { month: "short", day: "numeric" };
  return `${monday.toLocaleDateString(undefined, opts)}–${sunday.toLocaleDateString(undefined, opts)}, ${monday.getFullYear()}`;
};

const getDayTitle = (date: Date): string =>
  `${date.toLocaleString(undefined, { weekday: "long" })} ${date.getDate()} ${date.toLocaleString(undefined, { month: "long" })} ${date.getFullYear()}`;

/* -------------------------------- Event chip -------------------------------- */
const MonthEventChip: FC<{
  event: CalendarEvent;
  onClick: (id: string) => void;
}> = ({ event, onClick }) => {
  const styles = useStyles();
  const { bg, fg, border } = getTypeColors(event.type);
  const handleClick = (ev: MouseEvent<HTMLDivElement>): void => {
    ev.stopPropagation();
    onClick(event.id);
  };
  return (
    <div
      className={styles.chip}
      onClick={handleClick}
      style={{
        backgroundColor: bg,
        color: fg,
        border: border ? `1px solid ${border}` : undefined,
      }}
      title={`${formatTimeRange(event.start, event.end)} • ${event.title}`}
    >
      <span className={styles.chipTime}>{formatTimeRange(event.start, event.end)}</span>
      <span className={styles.chipTitle}>{event.title}</span>
    </div>
  );
};

/* --------------------------------- EventCard -------------------------------- */
const EventCard: FC<{
  item: PositionedEvent;
  onClick: (id: string) => void;
}> = ({ item, onClick }) => {
  const styles = useStyles();
  const { bg, fg, border } = getTypeColors(item.type);
  const isAppointment = item.type?.toLowerCase().includes("appointment");
  const titleColor = isAppointment ? "#000" : fg;
  
  // Calculate horizontal position with gap between lanes
  const laneGap = 2; // 2px gap between lanes
  const leftPct = (item.lane / Math.max(1, item.laneCount)) * 100;
  const widthPct = (100 / Math.max(1, item.laneCount)) - (laneGap / 10); // Subtract gap

  const handleClick = (ev: MouseEvent<HTMLDivElement>): void => {
    ev.stopPropagation();
    onClick(item.id);
  };

  return (
    <div
      className={styles.eventCard}
      onClick={handleClick}
      style={{
        top: `${item.topPct}%`,
        height: `${item.heightPct}%`,
        left: `${leftPct}%`,
        width: `calc(${widthPct}% - ${laneGap}px)`, // Proper width calculation with gap
        backgroundColor: bg,
        color: fg,
        border: border ? `1px solid ${border}` : undefined,
      }}
      title={`${formatTimeRange(item.displayStart, item.displayEnd)} • ${item.title}`}
    >
      <span className={styles.eventTitle} style={{ color: titleColor }}>
        {item.title}
      </span>
      {item.subtitle ? <span className={styles.eventSubtitle}>{item.subtitle}</span> : null}
    </div>
  );
};

/* ------------------------------ Header component ---------------------------- */
const HeaderBar: FC<{
  view: CalendarView;
  setView: (v: CalendarView) => void;
  focusDate: Date;
  setFocusDate: (d: Date) => void;
}> = ({ view, setView, focusDate, setFocusDate }) => {
  const styles = useStyles();

  const today = (): void => setFocusDate(new Date());
  const prev = (): void => {
    if (view === "month") setFocusDate(addDays(new Date(focusDate), -30));
    else if (view === "week") setFocusDate(addDays(new Date(focusDate), -7));
    else setFocusDate(addDays(new Date(focusDate), -1));
  };
  const next = (): void => {
    if (view === "month") setFocusDate(addDays(new Date(focusDate), 30));
    else if (view === "week") setFocusDate(addDays(new Date(focusDate), 7));
    else setFocusDate(addDays(new Date(focusDate), 1));
  };

  const periodTitle =
    view === "month"
      ? getMonthTitle(focusDate)
      : view === "week"
      ? getWeekTitle(addDays(getFirstMondayOfMatrix(focusDate), 0))
      : getDayTitle(focusDate);

  return (
    <div className={styles.header}>
      <div className={styles.headerLeft}>
        <Toolbar aria-label="calendar navigation">
          <ToolbarButton onClick={today} appearance="primary">Today</ToolbarButton>
          <ToolbarButton onClick={prev} aria-label="Previous">◀</ToolbarButton>
          <ToolbarButton onClick={next} aria-label="Next">▶</ToolbarButton>
        </Toolbar>
      </div>
      <div className={styles.headerCenter}>
        <Text size={500} weight="semibold" className={styles.periodTitle}>{periodTitle}</Text>
      </div>
      <div className={styles.headerRight}>
        <TabList selectedValue={view} onTabSelect={(_, data): void => setView(data.value as CalendarView)}>
          <Tab value="month">Month</Tab>
          <Tab value="week">Week</Tab>
          <Tab value="day">Day</Tab>
        </TabList>
      </div>
      <Divider />
    </div>
  );
};

/* ------------------------------- Month component ---------------------------- */
const MonthGrid: FC<{
  focusDate: Date;
  events: CalendarEvent[];
  onSelect: (id: string) => void;
}> = ({ focusDate, events, onSelect }) => {
  const styles = useStyles();
  const days = React.useMemo(() => buildMonthMatrix(focusDate), [focusDate]);
  const monthIdx = focusDate.getMonth();

  return (
    <div className={styles.monthGrid}>
      {days.map((day) => {
        const dayEvents = events.filter((e) => isSameDay(e.start, day));
        // responsive chip cap (approx; real cap enforced by CSS height)
        const visibleCap = 4;
        const showMore = dayEvents.length > visibleCap;
        const visible = dayEvents.slice(0, visibleCap);

        return (
          <div key={day.toISOString()} className={styles.dayCell}>
            <div className={styles.dayHeader}>
              <Text weight="semibold" className={day.getMonth() === monthIdx ? undefined : styles.outsideMonth}>
                {day.getDate()}
              </Text>
              <Badge appearance="filled" size="tiny">
                {dayEvents.length}
              </Badge>
            </div>
            {visible.map((ev) => (
              <MonthEventChip key={ev.id} event={ev} onClick={(id): void => onSelect(id)} />
            ))}
            {showMore ? <span className={styles.chipMore}>+{dayEvents.length - visibleCap} more</span> : null}
          </div>
        );
      })}
    </div>
  );
};

/* -------------------------------- Week component ---------------------------- */
const WeekGrid: FC<{
  focusDate: Date;
  events: CalendarEvent[];
  onSelect: (id: string) => void;
}> = ({ focusDate, events, onSelect }) => {
  const styles = useStyles();
  const monday = React.useMemo(() => {
    const firstMonday = getFirstMondayOfMatrix(focusDate);
    // Align monday to the week containing focusDate
    const deltaDays = Math.floor((startOfDay(focusDate).getTime() - firstMonday.getTime()) / DAY_MS);
    const normalizedMonday = addDays(firstMonday, Math.floor(deltaDays / 7) * 7);
    return normalizedMonday;
  }, [focusDate]);
  const days = React.useMemo(
    () => Array.from({ length: 7 }, (_, i) => addDays(monday, i)),
    [monday]
  );

  // Build left rail hour labels
  const hours = React.useMemo(() => Array.from({ length: 10 }, (_, i) => 9 + i), []);

  // Overlay hour/quarter lines positions (% from top of the timeline area, not including header)
  const hourLinesPct = React.useMemo(() => hours.map((h) => ((h - 9) / 9) * 100), [hours]);
  const quarterLinesPct = React.useMemo(
    () =>
      hours.flatMap((h) => {
        const base = ((h - 9) / 9) * 100;
        return [base + (100 / 9) * 0.25, base + (100 / 9) * 0.5, base + (100 / 9) * 0.75];
      }),
    [hours]
  );

  return (
    <div className={styles.schedulerRoot}>
      <div className={styles.leftRail} aria-hidden>
        {hours.slice(0, 9).map((h) => (
          <div key={h} className={styles.leftRailHour}>
            <Tooltip content={formatHourLabel(h)} relationship="label">
              <Text size={300}>{formatShortHourLabel(h)}</Text>
            </Tooltip>
          </div>
        ))}
      </div>
      <div className={styles.gridArea}>
        <div className={styles.weekColumns}>
          {/* Overlay hour/quarter lines across all columns */}
          <div className={styles.overlayLines}>
            {hourLinesPct.map((pct, idx) => (
              <div key={`h-${idx}`} className={styles.hourLine} style={{ top: `${pct}%` }} />
            ))}
            {quarterLinesPct.map((pct, idx) => (
              <div key={`q-${idx}`} className={styles.quarterLine} style={{ top: `${pct}%` }} />
            ))}
          </div>

          {days.map((day) => {
            const segments = computeDaySegments(day, events);
            const positioned = computeLanes(segments);
            const headerLabel = `${day.toLocaleString(undefined, { weekday: "short" })} ${day.getDate()}`;

            return (
              <div key={day.toISOString()} className={styles.dayColumn}>
                <div className={styles.dayColumnHeader}>
                  <Text weight="semibold">{headerLabel}</Text>
                </div>
                {/* Events container positioned below header */}
                <div className={styles.eventsContainer}>
                  {positioned.map((p) => (
                    <EventCard key={`${p.id}-${p.displayStart.toISOString()}`} item={p} onClick={(id): void => onSelect(id)} />
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};

/* --------------------------------- Day component ---------------------------- */
const DayTimeline: FC<{
  focusDate: Date;
  events: CalendarEvent[];
  onSelect: (id: string) => void;
}> = ({ focusDate, events, onSelect }) => {
  const styles = useStyles();
  const day = startOfDay(focusDate);
  const segments = React.useMemo(() => computeDaySegments(day, events), [day, events]);
  const positioned = React.useMemo(() => computeLanes(segments), [segments]);

  const hours = React.useMemo(() => Array.from({ length: 10 }, (_, i) => 9 + i), []);
  const hourLinesPct = React.useMemo(() => hours.map((h) => ((h - 9) / 9) * 100), [hours]);
  const quarterLinesPct = React.useMemo(
    () =>
      hours.flatMap((h) => {
        const base = ((h - 9) / 9) * 100;
        return [base + (100 / 9) * 0.25, base + (100 / 9) * 0.5, base + (100 / 9) * 0.75];
      }),
    [hours]
  );

  const headerLabel = `${day.toLocaleString(undefined, { weekday: "short" })} ${day.getDate()}`;

  return (
    <div className={styles.schedulerRoot}>
      <div className={styles.leftRail} aria-hidden>
        {hours.slice(0, 9).map((h) => (
          <div key={h} className={styles.leftRailHour}>
            <Text size={300}>{formatHourLabel(h)}</Text>
          </div>
        ))}
      </div>
      <div className={styles.gridArea} role="region" aria-label="Day schedule">
        <div className={styles.dayColumnWrapper}>
          {/* Day header */}
          <div className={styles.dayColumnHeader}>
            <Text weight="semibold">{headerLabel}</Text>
          </div>

          {/* Gridlines positioned below header */}
          <div className={styles.dayGridlines}>
            {hourLinesPct.map((pct, idx) => (
              <div key={`dh-${idx}`} className={styles.hourLine} style={{ top: `${pct}%` }} />
            ))}
            {quarterLinesPct.map((pct, idx) => (
              <div key={`dq-${idx}`} className={styles.quarterLine} style={{ top: `${pct}%` }} />
            ))}
          </div>

          {/* Events positioned below header */}
          <div className={styles.eventsContainer}>
            {positioned.map((p) => (
              <EventCard key={`${p.id}-${p.displayStart.toISOString()}`} item={p} onClick={(id): void => onSelect(id)} />
            ))}
          </div>
        </div>
      </div>
    </div>
  );
};

/* ---------------------------------- Main ------------------------------------ */
export const Calender: FC<CalendarProps> = (props) => {
  const {
    dataset,
    fromColumn,
    toColumn,
    titleColumn,
    subtitleColumn,
    typeColumn,
    onSelect,
    containerHeight,
  } = props;
  const styles = useStyles();
  const [view, setView] = React.useState<CalendarView>("month");
  const [focusDate, setFocusDate] = React.useState<Date>(new Date());

  const events = React.useMemo(
    () => normalizeEvents(dataset, titleColumn, subtitleColumn, typeColumn, fromColumn, toColumn),
    [dataset, titleColumn, subtitleColumn, typeColumn, fromColumn, toColumn]
  );

  const safeHeight = typeof containerHeight === "number" && containerHeight > 0 ? containerHeight : 800;

  const handleSelect = React.useCallback(
    (id: string): void => {
      if (onSelect) onSelect([id]);
    },
    [onSelect]
  );

  return (
    <FluentProvider theme={webLightTheme} style={{ height: safeHeight, width: "100%" }}>
      <div className={styles.root}>
        <HeaderBar view={view} setView={setView} focusDate={focusDate} setFocusDate={setFocusDate} />
        <div className={styles.content}>
          {view === "month" && <MonthGrid focusDate={focusDate} events={events} onSelect={handleSelect} />}
          {view === "week" && <WeekGrid focusDate={focusDate} events={events} onSelect={handleSelect} />}
          {view === "day" && <DayTimeline focusDate={focusDate} events={events} onSelect={handleSelect} />}
        </div>
      </div>
    </FluentProvider>
  );
};