import Constants from "expo-constants";
import * as Device from "expo-device";
import * as Linking from "expo-linking";
import * as Notifications from "expo-notifications";
import { StatusBar } from "expo-status-bar";
import { Component, useCallback, useEffect, useRef, useState } from "react";
import {
  ActivityIndicator,
  Alert,
  FlatList,
  InteractionManager,
  Platform,
  ScrollView,
  StyleSheet,
  Text,
  TextInput,
  TouchableOpacity,
  View,
} from "react-native";
import {
  DEFAULT_ORIGIN,
  apiBase,
  clearStoredSession,
  fetchInbox,
  fetchPinRoster,
  getStoredOrigin,
  getStoredToken,
  getStoredUser,
  loginPin,
  registerPushToken,
  setStoredOrigin,
  setStoredSession,
  technicianTerminalUrl,
  workOrderUrl,
} from "./src/api";

const BAKED_ORIGIN =
  String(Constants.expoConfig?.extra?.defaultApiOrigin || "").trim() || DEFAULT_ORIGIN;

const MAX_PIN = 6;
const ALLOWED_ROLES = ["artisan", "admin", "supervisor"];

function initials(label) {
  const parts = String(label || "").trim().split(/\s+/).filter(Boolean);
  if (!parts.length) return "?";
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

async function ensureAndroidChannel() {
  if (Platform.OS !== "android") return;
  try {
    await Notifications.setNotificationChannelAsync("breakdowns", {
      name: "Breakdowns & work orders",
      importance: Notifications.AndroidImportance.HIGH,
      vibrationPattern: [0, 250, 120, 250],
    });
  } catch {
    // Non-fatal — app must still open if channel setup fails.
  }
}

async function getFcmDeviceToken() {
  if (!Device.isDevice) return null;
  try {
    await ensureAndroidChannel();
    const { status: existing } = await Notifications.getPermissionsAsync();
    let status = existing;
    if (existing !== "granted") {
      const req = await Notifications.requestPermissionsAsync();
      status = req.status;
    }
    if (status !== "granted") return null;
    const tok = await Notifications.getDevicePushTokenAsync();
    return String(tok?.data || "").trim() || null;
  } catch {
    return null;
  }
}

function setupNotificationHandler() {
  try {
    Notifications.setNotificationHandler({
      handleNotification: async () => ({
        shouldShowAlert: true,
        shouldPlaySound: true,
        shouldSetBadge: true,
      }),
    });
  } catch {
    // Ignore — UI must still load.
  }
}

class ErrorBoundary extends Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error) {
    return { error };
  }

  render() {
    if (this.state.error) {
      return (
        <View style={styles.centered}>
          <Text style={styles.brand}>IRONLOG Notify</Text>
          <Text style={styles.error}>Something went wrong starting the app.</Text>
          <Text style={styles.hint}>{String(this.state.error?.message || this.state.error)}</Text>
        </View>
      );
    }
    return this.props.children;
  }
}

function AppBody() {
  const [ready, setReady] = useState(false);
  const [bootError, setBootError] = useState("");
  const [showServerEdit, setShowServerEdit] = useState(false);
  const [originInput, setOriginInput] = useState(BAKED_ORIGIN);
  const [origin, setOrigin] = useState(BAKED_ORIGIN);
  const [sessionToken, setSessionToken] = useState("");
  const [user, setUser] = useState(null);
  const [roster, setRoster] = useState([]);
  const [selectedUsername, setSelectedUsername] = useState("");
  const [pinValue, setPinValue] = useState("");
  const [loginError, setLoginError] = useState("");
  const [busy, setBusy] = useState(false);
  const [pushToken, setPushToken] = useState("");
  const [pushStatus, setPushStatus] = useState("");
  const [inbox, setInbox] = useState([]);
  const responseListener = useRef(null);

  const openFromNotificationData = useCallback((data) => {
    const woId = data?.wo_id || data?.work_order_id;
    if (woId && origin) {
      Linking.openURL(workOrderUrl(origin, woId));
      return;
    }
    if (origin) Linking.openURL(technicianTerminalUrl(origin));
  }, [origin]);

  const refreshInbox = useCallback(async () => {
    if (!origin || !sessionToken) return;
    try {
      const res = await fetchInbox(origin, sessionToken, 25);
      setInbox(Array.isArray(res.items) ? res.items : []);
    } catch {
      setInbox([]);
    }
  }, [origin, sessionToken]);

  const registerPush = useCallback(async (authToken, apiOrigin) => {
    try {
      const deviceToken = await getFcmDeviceToken();
      if (!deviceToken) {
        setPushStatus("Notification permission denied or unavailable on this device.");
        return;
      }
      setPushToken(deviceToken);
      const label = `${Device.modelName || "Android"} · ${Device.osVersion || ""}`.trim();
      const res = await registerPushToken(apiOrigin, authToken, deviceToken, label);
      setPushStatus(
        res.push_enabled
          ? "Registered for push alerts."
          : "App registered; server push pending Firebase setup."
      );
    } catch (err) {
      setPushStatus(err?.message || "Could not register push token.");
    }
  }, []);

  const loadRoster = useCallback(async (apiOrigin) => {
    try {
      const rosterRes = await fetchPinRoster(apiOrigin);
      setRoster(Array.isArray(rosterRes.technicians) ? rosterRes.technicians : []);
      setBootError("");
    } catch (err) {
      setRoster([]);
      setBootError(err?.message || "Could not reach IRONLOG server.");
    }
  }, []);

  useEffect(() => {
    let cancelled = false;

    async function boot() {
      try {
        setupNotificationHandler();
        const stored = await getStoredOrigin();
        const o = stored || BAKED_ORIGIN;
        if (!cancelled) {
          setOriginInput(o);
          setOrigin(o);
        }

        const tok = await getStoredToken();
        const u = await getStoredUser();

        if (tok && u) {
          if (!cancelled) {
            setSessionToken(tok);
            setUser(u);
          }
          InteractionManager.runAfterInteractions(() => {
            if (!cancelled) registerPush(tok, o).catch(() => {});
          });
        }

        await loadRoster(o);
      } catch (err) {
        if (!cancelled) setBootError(err?.message || "Startup failed.");
      } finally {
        if (!cancelled) setReady(true);
      }
    }

    boot();
    return () => {
      cancelled = true;
    };
  }, [loadRoster, registerPush]);

  useEffect(() => {
    if (!ready) return undefined;

    try {
      responseListener.current = Notifications.addNotificationResponseReceivedListener((response) => {
        const data = response?.notification?.request?.content?.data || {};
        openFromNotificationData(data);
      });
    } catch {
      // Non-fatal.
    }

    return () => {
      try {
        responseListener.current?.remove?.();
      } catch {
        // Ignore cleanup errors.
      }
      responseListener.current = null;
    };
  }, [ready, openFromNotificationData]);

  useEffect(() => {
    if (sessionToken) refreshInbox();
  }, [sessionToken, refreshInbox]);

  async function saveOrigin() {
    const o = await setStoredOrigin(originInput);
    if (!o) {
      Alert.alert("Server URL", "Enter a valid IRONLOG server address.");
      return;
    }
    setOrigin(o);
    setShowServerEdit(false);
    await loadRoster(o);
  }

  async function submitPinLogin() {
    if (!selectedUsername || pinValue.length < 4) return;
    setBusy(true);
    setLoginError("");
    try {
      const res = await loginPin(origin, selectedUsername, pinValue);
      const roles = Array.isArray(res.user?.roles) ? res.user.roles : [res.user?.role];
      const ok = roles.some((r) => ALLOWED_ROLES.includes(String(r).toLowerCase()));
      if (!ok) throw new Error("This app is for workshop technicians.");
      await setStoredSession({ token: res.token, user: res.user });
      setSessionToken(res.token);
      setUser(res.user);
      setPinValue("");
      await registerPush(res.token, origin);
      await refreshInbox();
    } catch (err) {
      setLoginError(err?.message || "Sign-in failed");
      setPinValue("");
    } finally {
      setBusy(false);
    }
  }

  async function signOut() {
    await clearStoredSession();
    setSessionToken("");
    setUser(null);
    setPushToken("");
    setPushStatus("");
    setInbox([]);
    setPinValue("");
    setSelectedUsername("");
    await loadRoster(origin);
  }

  function onPinDigit(d) {
    if (pinValue.length >= MAX_PIN) return;
    setPinValue((v) => v + d);
  }

  if (!ready) {
    return (
      <View style={styles.centered}>
        <ActivityIndicator size="large" color="#2563eb" />
        <Text style={styles.hint}>Loading IRONLOG Notify…</Text>
      </View>
    );
  }

  if (!sessionToken) {
    return (
      <ScrollView contentContainerStyle={styles.loginScroll} keyboardShouldPersistTaps="handled">
        <StatusBar style="light" />
        <Text style={styles.brand}>IRONLOG Notify</Text>
        <Text style={styles.sub}>Breakdown & work order alerts for technicians</Text>

        <Text style={styles.label}>IRONLOG server</Text>
        {showServerEdit ? (
          <View style={styles.row}>
            <TextInput
              style={styles.input}
              value={originInput}
              onChangeText={setOriginInput}
              autoCapitalize="none"
              autoCorrect={false}
              placeholder="https://ironlog.ironlogafrica.com"
              placeholderTextColor="#64748b"
            />
            <TouchableOpacity style={styles.btn} onPress={saveOrigin}>
              <Text style={styles.btnText}>Save</Text>
            </TouchableOpacity>
          </View>
        ) : (
          <View>
            <Text style={styles.serverValue}>{origin}</Text>
            <TouchableOpacity onPress={() => setShowServerEdit(true)}>
              <Text style={styles.link}>Change server</Text>
            </TouchableOpacity>
          </View>
        )}
        <Text style={styles.hint}>API: {apiBase(origin)}</Text>
        {bootError ? <Text style={styles.error}>{bootError}</Text> : null}

        <Text style={[styles.label, { marginTop: 20 }]}>Tap your name</Text>
        <View style={styles.roster}>
          {roster.map((u) => {
            const active = selectedUsername === u.username;
            return (
              <TouchableOpacity
                key={u.username}
                style={[styles.rosterBtn, active && styles.rosterBtnActive]}
                onPress={() => {
                  setSelectedUsername(u.username);
                  setPinValue("");
                  setLoginError("");
                }}
              >
                <View style={styles.avatar}>
                  <Text style={styles.avatarText}>{initials(u.label || u.username)}</Text>
                </View>
                <Text style={styles.rosterLabel}>{u.label || u.username}</Text>
              </TouchableOpacity>
            );
          })}
          {!roster.length ? <Text style={styles.hint}>No PIN users on server yet.</Text> : null}
        </View>

        <Text style={styles.pinDisplay}>{pinValue ? "•".repeat(pinValue.length) : "····"}</Text>
        <View style={styles.pad}>
          {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((n) => (
            <TouchableOpacity key={n} style={styles.padKey} onPress={() => onPinDigit(String(n))}>
              <Text style={styles.padKeyText}>{n}</Text>
            </TouchableOpacity>
          ))}
          <TouchableOpacity style={styles.padKeyMuted} onPress={() => setPinValue("")}>
            <Text style={styles.padKeyText}>Clr</Text>
          </TouchableOpacity>
          <TouchableOpacity style={styles.padKey} onPress={() => onPinDigit("0")}>
            <Text style={styles.padKeyText}>0</Text>
          </TouchableOpacity>
          <TouchableOpacity style={styles.padKeyMuted} onPress={() => setPinValue((v) => v.slice(0, -1))}>
            <Text style={styles.padKeyText}>⌫</Text>
          </TouchableOpacity>
        </View>

        <TouchableOpacity
          style={[styles.primaryBtn, (!selectedUsername || pinValue.length < 4 || busy) && styles.disabled]}
          disabled={!selectedUsername || pinValue.length < 4 || busy}
          onPress={submitPinLogin}
        >
          <Text style={styles.primaryBtnText}>{busy ? "Signing in…" : "Sign in with PIN"}</Text>
        </TouchableOpacity>
        {loginError ? <Text style={styles.error}>{loginError}</Text> : null}
      </ScrollView>
    );
  }

  const displayName = user?.full_name || user?.username || "Technician";

  return (
    <View style={styles.root}>
      <StatusBar style="light" />
      <View style={styles.header}>
        <View>
          <Text style={styles.headerTitle}>IRONLOG Notify</Text>
          <Text style={styles.headerSub}>{displayName}</Text>
        </View>
        <TouchableOpacity onPress={signOut}>
          <Text style={styles.link}>Sign out</Text>
        </TouchableOpacity>
      </View>

      <View style={styles.card}>
        <Text style={styles.cardTitle}>Push status</Text>
        <Text style={styles.hint}>{pushStatus || "Checking…"}</Text>
        {pushToken ? (
          <Text style={styles.mono} numberOfLines={1}>
            Token: {pushToken.slice(0, 24)}…
          </Text>
        ) : null}
        <TouchableOpacity style={styles.secondaryBtn} onPress={() => registerPush(sessionToken, origin)}>
          <Text style={styles.secondaryBtnText}>Re-register device</Text>
        </TouchableOpacity>
      </View>

      <View style={styles.card}>
        <Text style={styles.cardTitle}>Workshop</Text>
        <TouchableOpacity style={styles.primaryBtn} onPress={() => Linking.openURL(technicianTerminalUrl(origin))}>
          <Text style={styles.primaryBtnText}>Open technician terminal</Text>
        </TouchableOpacity>
      </View>

      <View style={[styles.card, styles.inboxCard]}>
        <View style={styles.rowBetween}>
          <Text style={styles.cardTitle}>Recent alerts</Text>
          <TouchableOpacity onPress={refreshInbox}>
            <Text style={styles.link}>Refresh</Text>
          </TouchableOpacity>
        </View>
        <FlatList
          style={styles.inboxList}
          data={inbox}
          keyExtractor={(item) => String(item.id)}
          ListEmptyComponent={<Text style={styles.hint}>No alerts yet.</Text>}
          renderItem={({ item }) => {
            const woId = item.data?.wo_id;
            return (
              <TouchableOpacity
                style={styles.inboxItem}
                onPress={() => {
                  if (woId) Linking.openURL(workOrderUrl(origin, woId));
                }}
              >
                <Text style={styles.inboxTitle}>{item.title}</Text>
                <Text style={styles.hint}>{item.body}</Text>
                <Text style={styles.inboxMeta}>{item.sent_at}</Text>
              </TouchableOpacity>
            );
          }}
        />
      </View>
    </View>
  );
}

export default function App() {
  return (
    <ErrorBoundary>
      <AppBody />
    </ErrorBoundary>
  );
}

const styles = StyleSheet.create({
  root: { flex: 1, backgroundColor: "#0f172a", padding: 16, paddingTop: 48 },
  centered: { flex: 1, alignItems: "center", justifyContent: "center", backgroundColor: "#0f172a", padding: 24 },
  loginScroll: { flexGrow: 1, backgroundColor: "#0f172a", padding: 20, paddingTop: 56 },
  brand: { color: "#f8fafc", fontSize: 28, fontWeight: "700" },
  sub: { color: "#94a3b8", marginTop: 6, marginBottom: 20 },
  label: { color: "#cbd5e1", fontSize: 13, fontWeight: "700", marginBottom: 6 },
  serverValue: { color: "#f8fafc", fontSize: 16, fontWeight: "700", marginBottom: 4 },
  hint: { color: "#94a3b8", fontSize: 13, marginTop: 4 },
  mono: { color: "#64748b", fontSize: 11, marginTop: 6 },
  row: { flexDirection: "row", marginBottom: 4 },
  rowBetween: { flexDirection: "row", justifyContent: "space-between", alignItems: "center" },
  input: {
    flex: 1,
    backgroundColor: "#1e293b",
    color: "#f8fafc",
    borderRadius: 10,
    paddingHorizontal: 12,
    paddingVertical: 10,
    borderWidth: 1,
    borderColor: "#334155",
    marginRight: 8,
  },
  btn: { backgroundColor: "#334155", borderRadius: 10, paddingHorizontal: 14, justifyContent: "center" },
  btnText: { color: "#f8fafc", fontWeight: "700" },
  roster: { flexDirection: "row", flexWrap: "wrap", marginHorizontal: -5 },
  rosterBtn: {
    width: "30%",
    minWidth: 96,
    backgroundColor: "#1e293b",
    borderRadius: 12,
    padding: 10,
    alignItems: "center",
    borderWidth: 1,
    borderColor: "#334155",
    margin: 5,
  },
  rosterBtnActive: { borderColor: "#2563eb", backgroundColor: "#172554" },
  avatar: {
    width: 44,
    height: 44,
    borderRadius: 22,
    backgroundColor: "#2563eb",
    alignItems: "center",
    justifyContent: "center",
    marginBottom: 6,
  },
  avatarText: { color: "#fff", fontWeight: "700" },
  rosterLabel: { color: "#e2e8f0", fontSize: 12, textAlign: "center" },
  pinDisplay: { color: "#f8fafc", fontSize: 28, letterSpacing: 8, textAlign: "center", marginVertical: 16 },
  pad: { flexDirection: "row", flexWrap: "wrap", justifyContent: "center", marginHorizontal: -4, marginBottom: 16 },
  padKey: {
    width: "28%",
    maxWidth: 110,
    backgroundColor: "#1e293b",
    borderRadius: 12,
    paddingVertical: 16,
    alignItems: "center",
    margin: 4,
  },
  padKeyMuted: {
    width: "28%",
    maxWidth: 110,
    backgroundColor: "#0f172a",
    borderRadius: 12,
    paddingVertical: 16,
    alignItems: "center",
    borderWidth: 1,
    borderColor: "#334155",
    margin: 4,
  },
  padKeyText: { color: "#f8fafc", fontSize: 20, fontWeight: "700" },
  primaryBtn: {
    backgroundColor: "#2563eb",
    borderRadius: 10,
    paddingVertical: 14,
    alignItems: "center",
    marginTop: 8,
  },
  primaryBtnText: { color: "#fff", fontWeight: "700" },
  secondaryBtn: {
    marginTop: 10,
    borderRadius: 10,
    paddingVertical: 10,
    alignItems: "center",
    borderWidth: 1,
    borderColor: "#334155",
  },
  secondaryBtnText: { color: "#93c5fd", fontWeight: "700" },
  disabled: { opacity: 0.5 },
  error: { color: "#fca5a5", marginTop: 10 },
  header: { flexDirection: "row", justifyContent: "space-between", alignItems: "center", marginBottom: 12 },
  headerTitle: { color: "#f8fafc", fontSize: 22, fontWeight: "700" },
  headerSub: { color: "#94a3b8", marginTop: 2 },
  link: { color: "#93c5fd", fontWeight: "700" },
  card: {
    backgroundColor: "#1e293b",
    borderRadius: 12,
    padding: 14,
    marginBottom: 12,
    borderWidth: 1,
    borderColor: "#334155",
  },
  inboxCard: { flex: 1, minHeight: 180 },
  inboxList: { flex: 1 },
  cardTitle: { color: "#f8fafc", fontSize: 16, fontWeight: "700", marginBottom: 4 },
  inboxItem: {
    borderTopWidth: 1,
    borderTopColor: "#334155",
    paddingVertical: 10,
  },
  inboxTitle: { color: "#f8fafc", fontWeight: "700" },
  inboxMeta: { color: "#64748b", fontSize: 11, marginTop: 4 },
});
