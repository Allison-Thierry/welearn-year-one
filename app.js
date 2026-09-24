"use strict";

const CONFIG = {
  learners: 11275,
  lessonsCompleted: 112466,
  minutesLearned: 806782,
  launchLessons: 40,
  currentLessons: 102,
  localLessons: 50,
  finalVideoSource: "",
};

const LEARNER_OPTIONS = [6420, 8340, 9650, 11275];

const COUNTRIES = [
  { id: "040", name: "Austria" },
  { id: "036", name: "Australia" },
  { id: "056", name: "Belgium" },
  { id: "076", name: "Brazil" },
  { id: "152", name: "Chile" },
  { id: "170", name: "Colombia" },
  { id: "188", name: "Costa Rica" },
  { id: "203", name: "Czech Republic" },
  { id: "222", name: "El Salvador" },
  { id: "246", name: "Finland" },
  { id: "250", name: "France" },
  { id: "276", name: "Germany" },
  { id: "320", name: "Guatemala" },
  { id: "372", name: "Ireland" },
  { id: "376", name: "Israel" },
  { id: "380", name: "Italy" },
  { id: "484", name: "Mexico" },
  { id: "528", name: "Netherlands" },
  { id: "591", name: "Panama" },
  { id: "604", name: "Peru" },
  { id: "616", name: "Poland" },
  { id: "620", name: "Portugal" },
  { id: "702", name: "Singapore" },
  { id: "724", name: "Spain" },
  { id: "756", name: "Switzerland" },
  { id: "826", name: "United Kingdom" },
  { id: "840", name: "United States" },
].sort((a, b) => a.name.localeCompare(b.name));

const COUNTRY_IDS = new Set(COUNTRIES.map((country) => country.id));
const COUNTRY_NAME_BY_ID = new Map(COUNTRIES.map((country) => [country.id, country.name]));

const TRANSLATIONS = [
  { language: "English", text: "Welcome to WeLearn", lang: "en", dir: "ltr" },
  { language: "Italian", text: "Benvenuti su WeLearn", lang: "it", dir: "ltr" },
  { language: "French", text: "Bienvenue sur WeLearn", lang: "fr", dir: "ltr" },
  { language: "German", text: "Willkommen bei WeLearn", lang: "de", dir: "ltr" },
  { language: "Polish", text: "Witamy w WeLearn", lang: "pl", dir: "ltr" },
  { language: "Spanish", text: "Bienvenidos a WeLearn", lang: "es", dir: "ltr" },
  { language: "Portuguese", text: "Bem-vindos ao WeLearn", lang: "pt", dir: "ltr" },
  { language: "Czech", text: "Vítejte ve WeLearn", lang: "cs", dir: "ltr" },
  { language: "Hebrew", text: "ברוכים הבאים ל-WeLearn", lang: "he", dir: "rtl" },
  { language: "Russian", text: "Добро пожаловать в WeLearn", lang: "ru", dir: "ltr" },
];

const COURSES = [
  "Effective Event Supervision",
  "Welcome to WeLearn!",
  "Sheffield - PIV",
  "Sheffield - District View",
  "Sheffield - Employees",
  "Sheffield - Scheduling Queue and Reschedules",
  "Sheffield - Store Details",
  "Planning a Pre Inventory Visit or Call",
  "Inventory Responsibilities",
  "Sheffield - Overview",
  "Prepare - Event Start",
  "Prepare - Overview",
  "Count - Controls",
  "Count - Area Ranges",
  "Count - Monitoring Performance",
  "Count - Alerts",
  "Count - Duplicate Areas",
  "Prepare - File Upload",
  "Count - Viewing reports",
  "Count - TCA",
  "Count - Voiding Areas",
  "Verify - Alterations",
  "Finish - Final Reports",
  "Finish - Evaluations",
  "Finish - Transmit Event",
  "RGIS Academy - Figure Laser",
  "RGIS Academy - What is Accuracy?",
  "RGIS Academy - Sense of Urgency",
  "RGIS Academy - MQ Counting Techniques",
  "RGIS Academy - AQ Counting Techniques",
  "RGIS Academy - Performance Tracker",
  "RGIS Academy - In Store Behaviour",
  "RGIS Academy - Career Opportunities",
  "RGIS Academy - Verticals",
  "Efficient Start up",
  "Efficient Wrap Up",
  "PIQA",
  "Accuracy Checks",
  "Flow and Deployment",
  "Sheffield - Schedule People",
  "Sheffield Advanced - Mass Scheduling",
  "Introduction to Merchandising",
  "Health & Safety in Merchandising",
  "Merchandising Basics - Auditor",
  "Merchandising Basics - Supervisor",
  "Placing Area Tickets",
  "WeLearn Basics: Quick Start Guide for Managers",
  "Mastering WeLearn: Advanced Functionalities for Instructors",
  "Mastering WeLearn: Advanced Functionalities for Content Creators",
  "The New Hire Guide",
  "Building Combo Habits on Your District",
  "Inventory Responsibilities Assignment",
  "Unlock Your Combo Power",
  "From Counter to Rising",
  "AI Basics 1 — Understanding AI & Responsible Use",
  "AI Basics 2 — Asking Better Questions and AI at RGIS",
  "Working with the Flow Plan Advisor",
  "RGIS Definition: AAV",
  "Understanding KPIs – Part 1: Inventory Performance",
  "Tableau - Introduction",
  "Event Health Check",
  "Helping Others Perform at Their Best: A Guide to On-the-floor Performance Coaching",
  "Operational Risks",
  "Speed Run — Books",
  "Tableau - APICS",
  "Understanding KPIs – Part 2: Business Performance",
  "Counting Tips: Line-by-Line Counting",
  "Counting Tips: Wide vs Narrow",
  "Understanding KPIs: Inventory Performance",
  "ESL Series 1 – Introduction to ESL Operations",
  "ESL Series 2 – Productivity & Teamwork",
  "The Buddy System",
  "ESL Series 3 – Leading an ESL Installation",
  "ESL Series 4 – Customer Communication",
  "Driver Safety Essentials",
  "ESL Series 5 – Productivity Management",
  "ESL Series 6 – Coaching & Supporting Teams",
  "Results Framework",
  "Delivering Outstanding Sales Leadership",
  "CNH Cycle Count – Process Overview for Field Teams",
  "Social Styles",
  "Questioning Skills",
  "Listening Techniques",
  "Negotiation",
  "CNH Cycle Count – Part 2: After the Plant Analysis",
  "Challenger Sales Model",
  "High Performance",
  "Objection Handling",
  "Pipeline",
  "New Business Approach",
  "Proposal Writing",
  "Account Plan",
  "Account Reviews",
  "Key Account Management – Relationship Mapping",
  "Parry's CRM - Leads",
  "Parry's CRM - Opportunities",
  "Parry's CRM - Sales Account",
  "Supply Chain Audit Services & Platform",
  "AI Intermediate 1 — Getting More from AI",
  "AI Intermediate 2 — AI in Action",
  "AI Intermediate 3 — Inside the ChatRGIS Workspace",
  "AI Intermediate 4 — Build Your AI Colleague",
];

// Excel completion dates in the same source order as COURSES. The launch view
// includes every lesson completed on or before 7 October 2025 (40 lessons).
const COURSE_COMPLETION_SERIALS = [45917,45930,45911,45910,45919,45922,45912,45923,45925,45910,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45937,45931,45932,45933,45933,45936,45937,45938,45939,45939,45943,45943,45950,45971,45971,45971,45960,45992,46010,46010,46030,46073,46073,46120,46120,46125,46128,46127,46134,46175,46206,46224,46232,46241,46241,46252,46260,46262,46267,46267,46268,46269,46273,46280,45939,45939,46049,46064,46078,46078,46078,46078,46065,46065,46065,46108,46108,46108,46122,46122,46122,46188,46188,46223,46238,46288,46288,46288,46288];

const CATALOGUE = COURSES
  .map((title, assetIndex) => ({ title, assetIndex, completedSerial: COURSE_COMPLETION_SERIALS[assetIndex] }))
  .sort((first, second) => first.completedSerial - second.completedSerial || first.assetIndex - second.assetIndex);

const SPOTLIGHTS = [
  {
    label: "The Globetrotter 🌍",
    course: "RGIS Academy – In Store Behaviour",
    stat: "22 countries",
    copy: "One lesson. A lot of passport stamps.",
    image: "assets/catalog/card-32.webp",
  },
  {
    label: "The Finisher 🏁",
    course: "Flow and Deployment",
    stat: "81.5% made it to the finish line.",
    copy: "824 completions from 1,011 learners.",
    image: "assets/catalog/card-39.webp",
  },
  {
    label: "The Breakout ⚡",
    course: "Unlock Your Combo Power",
    stat: "2,693+ completions",
    copy: "Across 17 countries. Not bad for one of the new kids.",
    image: "assets/catalog/card-53.webp",
  },
  { label: "Nicola's Favorite ❤️", placeholder: true },
  { label: "Patrick's Favorite ❤️", placeholder: true },
  { label: "Kris's Favorite ❤️", placeholder: true },
  { label: "Jennifer Sarmiento's Favorite ❤️", placeholder: true },
];

const state = {
  unlockedStage: 0,
  learnerComplete: false,
  mapCompleteOnce: false,
  mapGameComplete: false,
  mapSelected: new Set(),
  translationIndex: 0,
  learningComplete: false,
  catalogueComplete: false,
  catalogueStarted: false,
  spotlightOpened: new Set(),
  videoStarted: false,
  jenniferStarted: false,
  lineHeight: 0,
  pulseTimer: null,
};

const reduceMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
const numberFormatter = new Intl.NumberFormat("en-GB");
const $ = (selector, root = document) => root.querySelector(selector);
const $$ = (selector, root = document) => [...root.querySelectorAll(selector)];

function flashLightCue(chapter) {
  if (!chapter || reduceMotion) return;
  const node = $(".timeline-node", chapter);
  if (!node) return;

  chapter.classList.remove("light-flicker");
  void chapter.offsetWidth;
  chapter.classList.add("light-flicker");

  const bounds = node.getBoundingClientRect();
  const originX = bounds.left + bounds.width / 2;
  const originY = bounds.top + bounds.height / 2;
  for (let index = 0; index < 7; index += 1) {
    const spark = document.createElement("i");
    const angle = (Math.PI * 2 * index) / 7;
    const distance = 30 + (index % 3) * 12;
    spark.className = "light-spark";
    spark.style.left = `${originX}px`;
    spark.style.top = `${originY}px`;
    spark.style.setProperty("--spark-x", `${Math.cos(angle) * distance}px`);
    spark.style.setProperty("--spark-y", `${Math.sin(angle) * distance}px`);
    spark.style.animationDelay = `${index * 26}ms`;
    document.body.append(spark);
    window.setTimeout(() => spark.remove(), 820);
  }

  window.setTimeout(() => chapter.classList.remove("light-flicker"), 820);
}

function setUnlocked(stage, animatePulse = true) {
  if (stage <= state.unlockedStage) return;
  state.unlockedStage = stage;

  $$(".chapter").forEach((chapter) => {
    const chapterStage = Number(chapter.dataset.stage);
    const unlocked = chapterStage <= stage;
    chapter.classList.toggle("is-locked", !unlocked);
    chapter.classList.toggle("is-unlocked", unlocked);
    chapter.classList.toggle("is-current", chapterStage === stage);
    chapter.setAttribute("aria-hidden", String(!unlocked));
  });

  const newlyUnlocked = $(`.chapter[data-stage="${stage}"]`);
  if (newlyUnlocked) {
    newlyUnlocked.classList.remove("just-unlocked");
    requestAnimationFrame(() => newlyUnlocked.classList.add("just-unlocked"));
    window.setTimeout(() => newlyUnlocked.classList.remove("just-unlocked"), 1050);
  }

  document.body.classList.remove("has-unlock-flash");
  requestAnimationFrame(() => document.body.classList.add("has-unlock-flash"));
  window.setTimeout(() => document.body.classList.remove("has-unlock-flash"), 900);

  const completedChapter = $(`.chapter[data-stage="${Math.max(1, stage - 1)}"]`);
  flashLightCue(completedChapter);

  requestAnimationFrame(() => updateTimeline(animatePulse));
}

function updateTimeline(animatePulse = false) {
  if (!state.unlockedStage) return;
  let targetHeight;
  if (document.body.classList.contains("is-celebrating")) {
    targetHeight = $("#journey").scrollHeight - 2;
  } else if (state.jenniferStarted) {
    const jenniferChapter = $("#jennifer");
    targetHeight = jenniferChapter.offsetTop + jenniferChapter.offsetHeight / 2;
  } else if (state.videoStarted) {
    const journeyTop = $("#journey").getBoundingClientRect().top;
    const videoBounds = $("#video-placeholder, .real-video").getBoundingClientRect();
    targetHeight = videoBounds.bottom - journeyTop;
  } else {
    const currentChapter = $(`.chapter[data-stage="${state.unlockedStage}"]`);
    if (!currentChapter) return;
    targetHeight = currentChapter.offsetTop + currentChapter.offsetHeight / 2;
  }
  const live = $("#timeline-live");
  const pulse = $("#timeline-pulse");

  if (animatePulse && !reduceMotion) {
    window.clearTimeout(state.pulseTimer);
    live.classList.remove("is-zapping");
    void live.offsetWidth;
    live.classList.add("is-zapping");
    pulse.style.top = `${state.lineHeight}px`;
    pulse.classList.add("is-moving");
    requestAnimationFrame(() => {
      pulse.style.top = `${targetHeight}px`;
    });
    state.pulseTimer = window.setTimeout(() => {
      pulse.classList.remove("is-moving");
      live.classList.remove("is-zapping");
    }, 850);
  }

  live.style.height = `${targetHeight}px`;
  state.lineHeight = targetHeight;
}

function lightOpening() {
  const switchButton = $("#bulb-switch");
  if (document.body.classList.contains("is-lit")) return;
  switchButton.classList.add("is-flickering");
  window.setTimeout(() => {
    document.body.classList.add("is-lit");
    launchOpeningConfetti();
    switchButton.classList.remove("is-flickering");
    setUnlocked(1);
  }, reduceMotion ? 30 : 820);
}

function launchOpeningConfetti() {
  if (reduceMotion) return;
  const stage = $("#opening-confetti");
  const playId = Date.now();
  stage.classList.add("is-active");
  $$(".confetti-burst", stage).forEach((burst, index) => {
    burst.src = `${burst.dataset.src}?play=${playId}-${index}`;
  });
  window.setTimeout(() => {
    stage.classList.remove("is-active");
    $$(".confetti-burst", stage).forEach((burst) => burst.removeAttribute("src"));
  }, 4750);
}

function buildLearnerOptions() {
  const container = $("#learner-options");
  LEARNER_OPTIONS.forEach((value) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "number-bubble";
    button.dataset.value = String(value);
    button.textContent = numberFormatter.format(value);
    button.setAttribute("aria-label", `${numberFormatter.format(value)} learners`);
    button.addEventListener("click", completeLearnerGuess);
    container.append(button);
  });
}

function completeLearnerGuess() {
  if (state.learnerComplete) return;
  state.learnerComplete = true;
  $$(".number-bubble").forEach((bubble) => {
    const isCorrect = Number(bubble.dataset.value) === CONFIG.learners;
    bubble.classList.add(isCorrect ? "is-correct" : "is-dimmed");
    bubble.disabled = true;
  });
  $("#learner-reveal").hidden = false;
  requestAnimationFrame(() => setUnlocked(2));
  window.setTimeout(() => {
    $("#learner-reveal").scrollIntoView({ behavior: reduceMotion ? "auto" : "smooth", block: "center" });
  }, reduceMotion ? 20 : 260);
}

function countryId(feature) {
  return String(feature.id).padStart(3, "0");
}

async function buildMap() {
  const status = $("#map-status");
  try {
    const response = await fetch("assets/vendor/countries-110m.json");
    if (!response.ok) throw new Error(`Map data failed with ${response.status}`);
    const topology = await response.json();
    const featureCollection = topojson.feature(topology, topology.objects.countries);
    const mapFeatures = featureCollection.features.filter((feature) => countryId(feature) !== "010");
    const visibleWorld = { type: "FeatureCollection", features: mapFeatures };
    const width = 960;
    const height = 480;
    const svg = d3.select("#world-map").attr("viewBox", `0 0 ${width} ${height}`);
    const projection = d3.geoNaturalEarth1().fitExtent([[18, 18], [width - 18, height - 18]], visibleWorld);
    projection
      .scale(projection.scale() * 1.15)
      .clipExtent([[0, 0], [width, height]]);
    const path = d3.geoPath(projection);

    svg.append("path")
      .datum({ type: "Sphere" })
      .attr("d", path)
      .attr("fill", "#0b1c22");

    svg.append("g")
      .selectAll("path")
      .data(mapFeatures)
      .join("path")
      .attr("class", "country")
      .attr("data-country-id", (feature) => countryId(feature))
      .attr("d", path)
      .on("click", (_event, feature) => handleCountryClick(feature));

    const targetFeatures = mapFeatures.filter((feature) => COUNTRY_IDS.has(countryId(feature)));
    const manualTargetPoints = [
      { id: "702", properties: { name: "Singapore" }, coordinates: [103.8198, 1.3521] },
    ];
    const mapTargets = [...targetFeatures, ...manualTargetPoints];
    const generousTargets = new Set(["040", "056", "188", "203", "222", "372", "376", "528", "591", "620", "702", "756"]);
    const mediumTargets = new Set(["246", "250", "276", "380", "616", "724", "826"]);

    svg.append("g")
      .selectAll("circle")
      .data(mapTargets)
      .join("circle")
      .attr("class", "country-hotspot")
      .attr("data-country-id", (feature) => countryId(feature))
      .attr("cx", (feature) => feature.coordinates ? projection(feature.coordinates)[0] : path.centroid(feature)[0])
      .attr("cy", (feature) => feature.coordinates ? projection(feature.coordinates)[1] : path.centroid(feature)[1])
      .attr("r", 3.5);

    svg.append("g")
      .selectAll("circle")
      .data(mapTargets)
      .join("circle")
      .attr("class", "country-hit-target")
      .attr("data-country-id", (feature) => countryId(feature))
      .attr("cx", (feature) => feature.coordinates ? projection(feature.coordinates)[0] : path.centroid(feature)[0])
      .attr("cy", (feature) => feature.coordinates ? projection(feature.coordinates)[1] : path.centroid(feature)[1])
      .attr("r", (feature) => generousTargets.has(countryId(feature)) ? 12 : mediumTargets.has(countryId(feature)) ? 9 : 7)
      .on("click", (_event, feature) => handleCountryClick(feature));

    status.textContent = "Choose a country to begin.";
  } catch (error) {
    console.error(error);
    status.textContent = "The map could not load. Refresh to try again.";
  }
}

function setCountryClass(id, className, enabled = true) {
  d3.selectAll(`[data-country-id="${id}"]`).classed(className, enabled);
}

function handleCountryClick(feature) {
  if (state.mapGameComplete || state.unlockedStage < 2) return;
  const id = countryId(feature);
  const name = COUNTRY_NAME_BY_ID.get(id) || feature.properties.name;

  if (COUNTRY_IDS.has(id)) {
    if (state.mapSelected.has(id)) return;
    state.mapSelected.add(id);
    setCountryClass(id, "is-correct");
    $("#map-progress").textContent = `${state.mapSelected.size} / ?`;
    $("#map-status").textContent = `${name} — lit.`;
    if (state.mapSelected.size === COUNTRIES.length) completeMap(true);
    return;
  }

  state.mapGameComplete = true;
  setCountryClass(id, "is-wrong");
  $("#map-status").textContent = `${name} isn't on the list — revealing the rest…`;
  window.setTimeout(() => completeMap(false), reduceMotion ? 50 : 720);
}

function completeMap(perfect) {
  state.mapGameComplete = true;
  COUNTRY_IDS.forEach((id) => setCountryClass(id, "is-revealed"));
  $("#map-progress").textContent = "27 / 27";
  $("#map-status").textContent = "All 27 countries are lit.";
  $("#map-result-title").textContent = perfect ? "27/27 — Good job!" : "27 countries";
  $("#map-result-special").textContent = perfect ? "Okay, geography skills noted. 👀" : "";
  $("#map-result").hidden = false;
  $("#country-list").hidden = false;

  if (!state.mapCompleteOnce) {
    state.mapCompleteOnce = true;
    requestAnimationFrame(() => setUnlocked(3));
  } else {
    requestAnimationFrame(() => updateTimeline(false));
  }
}

function resetMap() {
  state.mapGameComplete = false;
  state.mapSelected.clear();
  d3.selectAll(".country, .country-hotspot, .country-hit-target")
    .classed("is-correct", false)
    .classed("is-revealed", false)
    .classed("is-wrong", false);
  $("#map-progress").textContent = "0 / ?";
  $("#map-status").textContent = "Fresh map. Go again.";
  $("#map-result").hidden = true;
  $("#country-list").hidden = false;
}

function buildCountryList() {
  const list = $("#country-list");
  COUNTRIES.forEach((country) => {
    const item = document.createElement("li");
    item.textContent = country.name;
    list.append(item);
  });
}

function addTranslationEcho(translation, index) {
  const echo = document.createElement("span");
  const positions = [
    [4, 12], [62, 14], [6, 76], [67, 72], [18, 34], [72, 40], [32, 84], [44, 8], [3, 50],
  ];
  const [left, top] = positions[index % positions.length];
  echo.className = "translation-echo";
  echo.textContent = translation.text;
  echo.lang = translation.lang;
  echo.dir = translation.dir;
  echo.style.left = `${left}%`;
  echo.style.top = `${top}%`;
  $("#translation-echoes").append(echo);
}

function advanceTranslation() {
  if (state.translationIndex >= TRANSLATIONS.length - 1) return;
  const phrase = $("#translation-phrase");
  addTranslationEcho(TRANSLATIONS[state.translationIndex], state.translationIndex);
  state.translationIndex += 1;
  const next = TRANSLATIONS[state.translationIndex];
  phrase.classList.add("is-changing");

  window.setTimeout(() => {
    phrase.textContent = next.text;
    phrase.lang = next.lang;
    phrase.dir = next.dir;
    $("#translation-language").textContent = next.language;
    $("#translation-count").textContent = `${String(state.translationIndex + 1).padStart(2, "0")} / 10`;
    phrase.classList.remove("is-changing");

    if (state.translationIndex === TRANSLATIONS.length - 1) {
      const button = $("#translate-button");
      button.disabled = true;
      $(".translation-stage").classList.add("is-final");
      window.setTimeout(() => {
        $("#language-reveal").hidden = false;
        setUnlocked(4);
      }, reduceMotion ? 20 : 420);
    }
  }, reduceMotion ? 0 : 180);
}

let holdAnimationFrame = null;
let holdStartTime = 0;
let holdValue = 0;
let holding = false;

function calculateGuess(elapsed) {
  return Math.min(250000, Math.round(elapsed * 12 + elapsed * elapsed * 0.015));
}

function updateHoldGuess(now) {
  if (!holding) return;
  holdValue = calculateGuess(now - holdStartTime);
  $("#guess-number").textContent = numberFormatter.format(holdValue);
  holdAnimationFrame = requestAnimationFrame(updateHoldGuess);
}

function startHolding(event) {
  if (state.learningComplete || holding) return;
  if (event.type === "keydown" && !["Enter", " "].includes(event.key)) return;
  event.preventDefault();
  holding = true;
  holdStartTime = performance.now();
  $("#hold-button").classList.add("is-holding");
  if (event.pointerId !== undefined) $("#hold-button").setPointerCapture(event.pointerId);
  holdAnimationFrame = requestAnimationFrame(updateHoldGuess);
}

function stopHolding(event) {
  if (!holding || state.learningComplete) return;
  if (event.type === "keyup" && !["Enter", " "].includes(event.key)) return;
  event.preventDefault();
  holding = false;
  cancelAnimationFrame(holdAnimationFrame);
  holdValue = Math.max(1100, holdValue);
  $("#guess-number").textContent = numberFormatter.format(holdValue);
  $("#hold-button").classList.remove("is-holding");
  completeLearningGuess();
}

function completeLearningGuess() {
  state.learningComplete = true;
  $("#hold-button").disabled = true;
  window.setTimeout(() => {
    $(".hold-stage").hidden = true;
    $("#learning-result").hidden = false;
    animateNumber($("#minutes-counter"), CONFIG.minutesLearned, reduceMotion ? 1 : 1600);
    setUnlocked(5);
  }, reduceMotion ? 20 : 520);
}

function animateNumber(element, target, duration) {
  const start = performance.now();
  function frame(now) {
    const progress = Math.min(1, (now - start) / duration);
    const eased = 1 - Math.pow(1 - progress, 3);
    element.textContent = numberFormatter.format(Math.round(target * eased));
    if (progress < 1) requestAnimationFrame(frame);
  }
  requestAnimationFrame(frame);
}

function cardAsset(index) {
  return `assets/catalog/card-${String(index + 1).padStart(2, "0")}.webp`;
}

function buildCatalogue() {
  const wall = $("#catalogue-wall");
  CATALOGUE.forEach((course, index) => {
    const card = document.createElement("figure");
    card.className = "catalogue-card";
    card.dataset.index = String(index);
    card.dataset.completedSerial = String(course.completedSerial);
    card.title = course.title;
    card.hidden = index >= CONFIG.launchLessons;
    const image = document.createElement("img");
    image.src = cardAsset(course.assetIndex);
    image.alt = course.title;
    image.loading = index < 15 ? "eager" : "lazy";
    image.decoding = "async";
    card.append(image);
    wall.append(card);
  });
}

function updateCatalogue(event) {
  const count = Number(event.currentTarget.value);
  const previousCount = Number($("#catalogue-live-count").dataset.count || CONFIG.launchLessons);
  const wall = $("#catalogue-wall");
  const windowElement = $(".catalogue-window");
  $("#catalogue-live-count").textContent = `${count} lessons`;
  $("#catalogue-live-count").dataset.count = String(count);
  event.currentTarget.setAttribute("aria-valuetext", `${count} Global Package lessons`);

  if (!state.catalogueStarted && count !== CONFIG.launchLessons) {
    state.catalogueStarted = true;
    $("#catalogue-era-overlay").classList.add("is-dismissed");
  }

  $$(".catalogue-card").forEach((card, index) => {
    const shouldShow = index < count;
    card.hidden = !shouldShow;
    if (shouldShow && index >= previousCount) {
      card.style.animationDelay = `${Math.min(220, (index - previousCount) * 22)}ms`;
      card.classList.add("is-new");
      window.setTimeout(() => {
        card.classList.remove("is-new");
        card.style.animationDelay = "";
      }, 1100);
    }
  });

  const catalogueProgress = (count - CONFIG.launchLessons) / (CONFIG.currentLessons - CONFIG.launchLessons);
  requestAnimationFrame(() => {
    const scrollRange = Math.max(0, wall.scrollHeight - wall.clientHeight);
    wall.scrollTo({ top: scrollRange * catalogueProgress, behavior: "auto" });
  });

  if (count < CONFIG.currentLessons) {
    windowElement.classList.remove("is-at-end");
  }

  if (count === CONFIG.currentLessons && !state.catalogueComplete) {
    state.catalogueComplete = true;
    $("#catalogue-reveal").hidden = false;
    window.setTimeout(() => {
      wall.scrollTo({ top: wall.scrollHeight, behavior: reduceMotion ? "auto" : "smooth" });
      windowElement.classList.add("is-at-end");
    }, 260);
    requestAnimationFrame(() => setUnlocked(6));
  }
}

function buildSpotlights() {
  const grid = $("#spotlight-grid");
  SPOTLIGHTS.forEach((spotlight, index) => {
    const card = document.createElement("button");
    card.type = "button";
    card.className = `spotlight-card${spotlight.placeholder ? " is-placeholder" : ""}`;
    card.setAttribute("aria-expanded", "false");
    card.setAttribute("aria-label", `Reveal ${spotlight.label}`);
    card.innerHTML = `
      <span class="spotlight-card-inner">
        <span class="spotlight-face spotlight-front">
          <span>Flip to reveal</span>
          <strong>${spotlight.label}</strong>
        </span>
        <span class="spotlight-face spotlight-back">
          ${spotlight.image ? `<img src="${spotlight.image}" alt="" loading="lazy">` : ""}
          <span class="spotlight-kicker">${spotlight.label}</span>
          <strong class="spotlight-course">${spotlight.placeholder ? "Team pick" : spotlight.course}</strong>
          <span class="spotlight-stat">${spotlight.placeholder ? "Coming soon…" : spotlight.stat}</span>
          <span class="spotlight-copy">${spotlight.placeholder ? "Top secret for now." : spotlight.copy}</span>
        </span>
      </span>`;
    card.addEventListener("click", () => revealSpotlight(card, index));
    grid.append(card);
  });
}

function revealSpotlight(card, index) {
  if (state.spotlightOpened.has(index)) return;
  state.spotlightOpened.add(index);
  card.classList.add("is-flipped");
  card.setAttribute("aria-expanded", "true");
  if (state.spotlightOpened.size === 1) {
    $("#spotlight-hint").textContent = "The final stretch is lit. The other cards are still yours to explore.";
    window.setTimeout(() => setUnlocked(7), reduceMotion ? 20 : 420);
  }
}

function startVideoPlaceholder() {
  const shell = $("#video-shell");
  state.videoStarted = true;
  shell.hidden = false;
  shell.classList.add("is-playing");
  $("#video-cta").disabled = true;
  requestAnimationFrame(() => updateTimeline(true));
  window.setTimeout(() => shell.scrollIntoView({ behavior: reduceMotion ? "auto" : "smooth", block: "center" }), 60);

  if (CONFIG.finalVideoSource) {
    const video = document.createElement("video");
    video.controls = true;
    video.autoplay = true;
    video.src = CONFIG.finalVideoSource;
    video.className = "real-video";
    video.addEventListener("ended", revealJenniferMoment);
    $("#video-placeholder").replaceWith(video);
  }
}

function revealJenniferMoment() {
  if (state.jenniferStarted) return;
  state.jenniferStarted = true;
  $("#simulate-end").disabled = true;
  $("#jennifer").hidden = false;
  document.body.classList.add("is-holding-thought");
  setUnlocked(8);
  window.setTimeout(() => {
    $("#jennifer").scrollIntoView({ behavior: reduceMotion ? "auto" : "smooth", block: "center" });
  }, 120);
  window.setTimeout(() => {
    $("#jennifer-continue").disabled = false;
  }, reduceMotion ? 20 : 8300);
}

function revealFinale() {
  if (document.body.classList.contains("is-celebrating")) return;
  document.body.classList.remove("is-holding-thought");
  document.body.classList.add("is-celebrating");
  $("#celebration").hidden = false;
  setUnlocked(9);
  window.setTimeout(() => {
    $("#celebration").scrollIntoView({ behavior: reduceMotion ? "auto" : "smooth", block: "center" });
  }, 180);
  window.setTimeout(buildConfetti, reduceMotion ? 20 : 620);
}

function buildConfetti() {
  const container = $("#confetti");
  if (container.children.length || reduceMotion) return;
  const colors = ["#f6cb69", "#dd233f", "#ffffff", "#b8aa9d", "#285f9f"];
  for (let index = 0; index < 76; index += 1) {
    const piece = document.createElement("i");
    piece.className = "confetti-piece";
    piece.style.left = `${(index * 37) % 101}%`;
    piece.style.background = colors[index % colors.length];
    piece.style.setProperty("--delay", `${(index % 13) * 0.06}s`);
    piece.style.setProperty("--duration", `${3.2 + (index % 9) * 0.18}s`);
    piece.style.setProperty("--drift", `${((index * 19) % 180) - 90}px`);
    piece.style.setProperty("--spin", `${360 + (index % 7) * 140}deg`);
    container.append(piece);
  }
}

function wireEvents() {
  $("#bulb-switch").addEventListener("click", lightOpening);
  $("#map-retry").addEventListener("click", resetMap);
  $("#translate-button").addEventListener("click", advanceTranslation);

  const holdButton = $("#hold-button");
  holdButton.addEventListener("pointerdown", startHolding);
  holdButton.addEventListener("pointerup", stopHolding);
  holdButton.addEventListener("pointercancel", stopHolding);
  holdButton.addEventListener("keydown", startHolding);
  holdButton.addEventListener("keyup", stopHolding);

  $("#catalogue-slider").addEventListener("input", updateCatalogue);
  $("#video-cta").addEventListener("click", startVideoPlaceholder);
  $("#simulate-end").addEventListener("click", revealJenniferMoment);
  $("#jennifer-continue").addEventListener("click", revealFinale);
  window.addEventListener("resize", () => updateTimeline(false));
  window.addEventListener("load", () => updateTimeline(false));
}

function initialise() {
  if (CATALOGUE.length !== CONFIG.currentLessons || COURSE_COMPLETION_SERIALS.length !== COURSES.length) {
    console.error(`Expected ${CONFIG.currentLessons} dated courses, found ${CATALOGUE.length}.`);
  }
  buildLearnerOptions();
  buildCountryList();
  buildCatalogue();
  buildSpotlights();
  wireEvents();
  buildMap();
}

initialise();
