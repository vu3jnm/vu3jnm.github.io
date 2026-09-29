from datetime import date, timedelta
from zoneinfo import ZoneInfo
import argparse
import csv
import json
from pathlib import Path

from astral import Observer
from astral.sun import sun

LATITUDE = 15.86
LONGITUDE = 74.51
TIMEZONE = "Asia/Kolkata"
YEAR = 2026


def calculate_days():
    observer = Observer(latitude=LATITUDE, longitude=LONGITUDE)
    local_tz = ZoneInfo(TIMEZONE)
    day = date(YEAR, 1, 1)
    end = date(YEAR + 1, 1, 1)
    rows = []

    while day < end:
        events = sun(observer, date=day, tzinfo=local_tz)
        rows.append({
            "date": day.isoformat(),
            "sunrise": events["sunrise"].strftime("%H:%M:%S"),
            "sunset": events["sunset"].strftime("%H:%M:%S"),
        })
        day += timedelta(days=1)

    return rows


def write_csv(path, rows):
    with path.open("w", newline="", encoding="utf-8") as file:
        writer = csv.writer(file)
        writer.writerow(["date", "sunrise_local", "sunset_local", "timezone"])
        for row in rows:
            writer.writerow([row["date"], row["sunrise"], row["sunset"], TIMEZONE])


def write_clock(path, rows):
    data = json.dumps(rows, separators=(",", ":"))
    page = r'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Sunrise & Sunset Clock — 2026</title>
<style>
:root { color-scheme: dark; font-family: system-ui, sans-serif; background: #101923; color: #f4f6f8; }
* { box-sizing: border-box; }
body { min-height: 100vh; margin: 0; display: grid; place-items: center; padding: 24px; background: radial-gradient(ellipse at 50% 0%, #26394a, #101923 65%); }
main { width: min(100%, 680px); text-align: center; }
.eyebrow { color: #a8bdca; letter-spacing: .16em; text-transform: uppercase; font-size: .78rem; }
#clock { margin: 14px 0 4px; font-size: clamp(3.6rem, 14vw, 7.8rem); line-height: 1; font-variant-numeric: tabular-nums; letter-spacing: .03em; }
#today { color: #c2d0d8; font-size: 1.08rem; }
.times { display: grid; grid-template-columns: 1fr 1fr; gap: 18px; margin-top: 44px; }
.event { padding: 22px 12px; border-top: 1px solid #526675; }
.label { color: #a8bdca; font-size: .82rem; letter-spacing: .12em; text-transform: uppercase; }
.value { margin-top: 9px; font-size: clamp(1.8rem, 7vw, 2.7rem); font-variant-numeric: tabular-nums; }
.countdown { margin-top: 6px; color: #c2d0d8; font-size: .92rem; font-variant-numeric: tabular-nums; }
.verses { margin-top: 16px; line-height: 1.8; font-size: .98rem; }
.read-button { width: auto; margin-top: 12px; padding: 8px 14px; font-size: .9rem; }
.note { margin-top: 28px; color: #a8bdca; font-size: .9rem; }
.alarm-controls { margin-top: 26px; }
.alarm-button { border: 1px solid #7890a1; border-radius: 999px; background: #26394a; color: #f4f6f8; padding: 11px 20px; font: inherit; cursor: pointer; }
.alarm-button:hover { background: #344c60; }
#alarm-status { min-height: 1.5em; margin-top: 10px; color: #a8bdca; font-size: .9rem; }
@media (max-width: 420px) { .times { gap: 4px; } }
</style>
</head>
<body>
<main>
  <div class="eyebrow">Local time · Asia/Kolkata</div>
  <div id="clock" role="timer" aria-live="off">--:--:--</div>
  <div id="today">Loading date…</div>
  <section class="times" aria-label="Today's sun times">
    <div class="event"><div class="label">Sunrise</div><div class="value" id="sunrise">—</div><div class="countdown" id="sunrise-countdown">—</div><div class="verses" lang="sa" id="sunrise-verses">सूर्याय स्वाहा| सूर्याय इदम् न मम<br>प्रजापतये स्वाहा| प्रजापतये इदम् न मम</div><button class="read-button" type="button" data-read="sunrise-verses">Read aloud</button></div>
    <div class="event"><div class="label">Sunset</div><div class="value" id="sunset">—</div><div class="countdown" id="sunset-countdown">—</div><div class="verses" lang="sa" id="sunset-verses">अग्नये स्वाहा | अग्नये इदम् न मम<br>प्रजापतये स्वाहा| प्रजापतये इदम् न मम</div><button class="read-button" type="button" data-read="sunset-verses">Read aloud</button></div>
  </section>
  <div class="alarm-controls">
    <button class="alarm-button" id="alarm-toggle" type="button">Enable sound alarms</button>
    <div id="alarm-status" aria-live="polite">Enable sound to hear alarms. Keep this page open.</div>
  </div>
  <div class="note" id="note">15.86° N, 74.51° E · 2026 almanac</div>
</main>
<script>
const sunTimes = __SUN_DATA__;
const formatter = new Intl.DateTimeFormat("en-IN", {
  timeZone: "Asia/Kolkata", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false
});
const dateFormatter = new Intl.DateTimeFormat("en-IN", {
  timeZone: "Asia/Kolkata", weekday: "long", day: "numeric", month: "long", year: "numeric"
});
document.querySelectorAll('[data-read]').forEach(button => {
  button.addEventListener('click', () => {
    if (!("speechSynthesis" in window)) {
      document.getElementById("alarm-status").textContent = "Speech output is not supported by this browser.";
      return;
    }
    window.speechSynthesis.cancel();
    const text = document.getElementById(button.dataset.read).innerText.replace(/\s*\n\s*/g, ". ");
    const utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = "sa-IN";
    const voices = window.speechSynthesis.getVoices();
    utterance.voice = voices.find(voice => voice.lang.toLowerCase().startsWith("sa"))
      || voices.find(voice => voice.lang.toLowerCase().startsWith("hi"))
      || null;
    utterance.onstart = () => document.getElementById("alarm-status").textContent = "Reading Sanskrit aloud…";
    utterance.onend = () => document.getElementById("alarm-status").textContent = "Finished reading.";
    utterance.onerror = () => document.getElementById("alarm-status").textContent = "Could not read aloud. Check that your browser has a suitable voice installed.";
    window.speechSynthesis.speak(utterance);
  });
});
let alarmsEnabled = false;
let audioContext = null;
const firedAlarms = new Set();

function beep(frequency, startDelay) {
  const start = audioContext.currentTime + startDelay;
  const oscillator = audioContext.createOscillator();
  const gain = audioContext.createGain();
  oscillator.type = "sine";
  oscillator.frequency.value = frequency;
  gain.gain.setValueAtTime(0.0001, start);
  gain.gain.exponentialRampToValueAtTime(0.22, start + 0.02);
  gain.gain.exponentialRampToValueAtTime(0.0001, start + 0.28);
  oscillator.connect(gain);
  gain.connect(audioContext.destination);
  oscillator.start(start);
  oscillator.stop(start + 0.3);
}

function ringAlarm(label, alarmKey) {
  if (!alarmsEnabled || firedAlarms.has(alarmKey)) return;
  firedAlarms.add(alarmKey);
  beep(880, 0);
  beep(660, 0.38);
  beep(880, 0.76);
  document.getElementById("alarm-status").textContent = `Alarm: ${label}`;
}

function checkAlarms(now, localDate, entry) {
  if (!alarmsEnabled || !entry) return;
  for (const event of [
    { name: "sunrise", time: entry.sunrise, title: "Sunrise" },
    { name: "sunset", time: entry.sunset, title: "Sunset" }
  ]) {
    const eventTime = new Date(`${localDate}T${event.time}+05:30`);
    const alarmTimes = [
      { at: eventTime.getTime() - 10 * 60 * 1000, label: `${event.title} in 10 minutes` },
      { at: eventTime.getTime(), label: event.title }
    ];
    for (const alarm of alarmTimes) {
      const key = `${localDate}-${event.name}-${alarm.at}`;
      if (now.getTime() >= alarm.at && now.getTime() < alarm.at + 1500) {
        ringAlarm(alarm.label, key);
      }
    }
  }
}

function nextEventCountdown(now, localDate, eventTime) {
  if (!eventTime) return "No 2026 data";
  const target = new Date(`${localDate}T${eventTime}+05:30`);
  if (now >= target) target.setTime(target.getTime() + 86400000);
  const totalSeconds = Math.max(0, Math.floor((target.getTime() - now.getTime()) / 1000));
  const days = Math.floor(totalSeconds / 86400);
  const hours = Math.floor((totalSeconds % 86400) / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  const time = [hours, minutes, seconds].map(value => String(value).padStart(2, "0")).join(":");
  return days ? `${days}d ${time} until next` : `${time} until next`;
}

document.getElementById("alarm-toggle").addEventListener("click", async () => {
  if (alarmsEnabled) {
    alarmsEnabled = false;
    document.getElementById("alarm-toggle").textContent = "Enable sound alarms";
    document.getElementById("alarm-status").textContent = "Sound alarms are off.";
    return;
  }
  const AudioContextClass = window.AudioContext || window.webkitAudioContext;
  if (!AudioContextClass) {
    document.getElementById("alarm-status").textContent = "This browser does not support sound alarms.";
    return;
  }
  audioContext = audioContext || new AudioContextClass();
  if (audioContext.state === "suspended") await audioContext.resume();
  alarmsEnabled = true;
  beep(660, 0);
  document.getElementById("alarm-toggle").textContent = "Turn sound alarms off";
  document.getElementById("alarm-status").textContent = "Sound alarms on: 10 minutes before sunrise and sunset, and at both events.";
});

function updateClock() {
  const now = new Date();
  document.getElementById("clock").textContent = formatter.format(now);
  document.getElementById("today").textContent = dateFormatter.format(now);
  const localDate = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Kolkata", year: "numeric", month: "2-digit", day: "2-digit"
  }).format(now);
  const entry = sunTimes.find(item => item.date === localDate);
  checkAlarms(now, localDate, entry);
  document.getElementById("sunrise").textContent = entry ? entry.sunrise : "—";
  document.getElementById("sunset").textContent = entry ? entry.sunset : "—";
  document.getElementById("sunrise-countdown").textContent = nextEventCountdown(now, localDate, entry && entry.sunrise);
  document.getElementById("sunset-countdown").textContent = nextEventCountdown(now, localDate, entry && entry.sunset);
  document.getElementById("note").textContent = entry
    ? "15.86° N, 74.51° E · 2026 almanac"
    : "Sunrise and sunset data is available for 2026 only · 15.86° N, 74.51° E";
}
updateClock();
setInterval(updateClock, 1000);
</script>
</body>
</html>
'''.replace("__SUN_DATA__", data)
    path.write_text(page, encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(
        description="Create 2026 sunrise/sunset data and a live local digital clock."
    )
    parser.add_argument("--output", default="sunrise_sunset_2026.csv", help="CSV output path")
    parser.add_argument("--clock", default="clock.html", help="Clock webpage path")
    args = parser.parse_args()

    rows = calculate_days()
    csv_path = Path(args.output)
    clock_path = Path(args.clock)
    write_csv(csv_path, rows)
    write_clock(clock_path, rows)
    print(f"Wrote daily sun times to {csv_path}")
    print(f"Wrote digital clock to {clock_path}")


if __name__ == "__main__":
    main()
