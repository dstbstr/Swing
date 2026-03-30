import SendMonthlyReport from "./Reports/MonthlyReport";
import EnsureThisMonth, { EnsureNextMonth, UpdateVolunteers } from "./TimerScripts/TimerScripts";
import CopyLatestWaiverToAttendance from "./WaiverWatcher/WaiverWatcher";

const gasGlobal = globalThis as unknown as Record<string, unknown>;

gasGlobal.__SendMonthlyReportImpl = SendMonthlyReport;
gasGlobal.__EnsureThisMonthImpl = EnsureThisMonth;
gasGlobal.__EnsureNextMonthImpl = EnsureNextMonth;
gasGlobal.__UpdateVolunteersImpl = UpdateVolunteers;
gasGlobal.__CopyLatestWaiverToAttendanceImpl = CopyLatestWaiverToAttendance;
