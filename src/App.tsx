import React, { useEffect, useState } from "react";
import {
  PublicClientApplication,
  AccountInfo,
  Configuration,
} from "@azure/msal-browser";

const msalConfig: Configuration = {
  auth: {
    clientId: process.env.REACT_APP_CLIENT_ID!,
    authority: `https://login.microsoftonline.com/${process.env.REACT_APP_TENANT_ID}`,
    redirectUri: "/", // Make sure this matches your app's root
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

const graphScopes = ["OnlineMeetings.Read"];
const msalInstance = new PublicClientApplication(msalConfig);

const App: React.FC = () => {
  const [username, setUsername] = useState<string>("");
  const [meetingUrl, setMeetingUrl] = useState<string>("");
  const [attendance, setAttendance] = useState<any[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string>("");
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [accessToken, setAccessToken] = useState<string>("");

  useEffect(() => {
    const init = async () => {
      await msalInstance.initialize();
      const existing = msalInstance.getAllAccounts()[0];
      if (existing) {
        setAccount(existing);
        setUsername(existing.username);
        try {
          const result = await msalInstance.acquireTokenSilent({
            scopes: graphScopes,
            account: existing,
          });
          setAccessToken(result.accessToken);
        } catch (e) {
          console.warn("Silent auth failed. User must sign in.");
        }
      }
    };

    init();
  }, []);

  const handleLogin = async () => {
    try {
      const result = await msalInstance.loginPopup({ scopes: graphScopes });
      setAccount(result.account!);
      setUsername(result.account?.username || "");

      const tokenResult = await msalInstance.acquireTokenSilent({
        scopes: graphScopes,
        account: result.account!,
      });
      setAccessToken(tokenResult.accessToken);
    } catch (err: any) {
      console.error("Login failed", err);
      setError(err.message || JSON.stringify(err));
    } 
   };

  const handleFetch = async () => {
    if (!account || !accessToken) {
      setError("Please sign in first.");
      return;
    }

    setLoading(true);
    setError("");
    setAttendance([]);

    try {
      const encodedJoinUrl = encodeURIComponent(meetingUrl.trim());
      const meetingRes = await fetch(
        `https://graph.microsoft.com/v1.0/me/onlineMeetings?$filter=JoinWebUrl eq '${encodedJoinUrl}'`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const meetingData = await meetingRes.json();
      const meeting = meetingData.value?.[0];

      if (!meeting) throw new Error("Meeting not found. Make sure the URL is correct and you have access.");

      const reportsRes = await fetch(
        `https://graph.microsoft.com/v1.0/me/onlineMeetings/${meeting.id}/attendanceReports`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const reports = await reportsRes.json();
      if (!reports.value?.length) throw new Error("No attendance reports found");

      const reportId = reports.value[0].id;

      const attendeesRes = await fetch(
        `https://graph.microsoft.com/v1.0/me/onlineMeetings/${meeting.id}/attendanceReports/${reportId}/attendanceRecords`,
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );

      const attendees = await attendeesRes.json();
      setAttendance(attendees.value || []);
    } catch (err: any) {
      console.error("Error fetching meeting data:", err);
      setError(err.message || "An unknown error occurred.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ fontFamily: "sans-serif", padding: "2rem", maxWidth: "600px", margin: "0 auto" }}>
      <h1>📊 Teams Attendance Viewer</h1>

      {!account ? (
        <button onClick={handleLogin}>🔐 Sign in to Microsoft</button>
      ) : (
        <p>✅ Signed in as <strong>{username}</strong></p>
      )}

      {account && (
        <>
          <input
            type="text"
            placeholder="Paste full Teams Join URL"
            value={meetingUrl}
            onChange={(e) => setMeetingUrl(e.target.value)}
            style={{ width: "100%", padding: "0.5rem", marginTop: "1rem" }}
          />
          <button
            onClick={handleFetch}
            style={{ marginTop: "0.5rem", padding: "0.5rem 1rem" }}
            disabled={!meetingUrl.trim()}
          >
            Fetch Attendance
          </button>
        </>
      )}

      {loading && <p>⏳ Loading attendance data...</p>}
      {error && <p style={{ color: "red" }}>❌ {error}</p>}

      {!loading && attendance.length > 0 && (
        <>
          <h2>📋 Attendance:</h2>
          <ul>
            {attendance.map((a, i) => (
              <li key={i}>
                {a.identity?.displayName || "Unknown"} —{" "}
                {Math.round((a.totalAttendanceInSeconds || 0) / 60)} min
              </li>
            ))}
          </ul>
        </>
      )}
    </div>
  );
};

export default App;
