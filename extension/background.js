/**
 * @name Preso extension
 *
 * This script, as part of a Chrome extension, allows the refreshing and looping
 * of Google Slides documents (without having to resort to "Publishing to web").
 *
 * See: http://plemont.github.io/javascript/slides-api/2016/12/08/dynamic-dashboarding-with-slides-api.html
 */
const REFRESH_MS = 60000;
const PRESO_REGEX = /^https:\/\/docs\.google\.com\/presentation\/d\/[^/]+\/present(.*)$/;

function fullscreenAndLoop(tab) {
  chrome.windows.getCurrent(win =>
    chrome.windows.update(win.id, {state: 'fullscreen'}));
  let nextUrl = calculateNextSlideUrl(tab.url);
  setTimeout(createReload(tab.id, nextUrl), REFRESH_MS);
}

function calculateNextSlideUrl(url) {
  let [hostPath, parts] = url.split('?');
  let params = extractParamsDictionary(parts);
  let slideId = params.slide;
  let matches;
  if (slideId) {
    let idRegex = /^(id\..*)_(\d+)_(\d+)$/;
    if ((matches = idRegex.exec(slideId)) !== null) {
      let currentPage = +matches[2];
      let totalPages = +matches[3];
      let newSlide = [matches[1], (currentPage + 1) % totalPages,
          totalPages].join('_');
      return hostPath + '?loop=1&slide=' + newSlide;
    }
  }
  return url;
}

function createReload(tabId, nextUrl) {
  return function() {
    chrome.tabs.query({
      active: true,
      lastFocusedWindow: true
    }, function(tabs) {
      let url = tabs[0].url;
      let matches = PRESO_REGEX.exec(url);
      if (matches) {
        chrome.tabs.update(tabId, {url: nextUrl});
      }
    });
  };
}

function extractParamsDictionary(parts) {
  let params = {};
  parts.split('&').forEach(part => {
    let [key, value] = part.split('=');
    params[key] = value;
  });
  return params;
}

function checkForValidUrl(tabId, changeInfo, tab) {
  // Only process events that are completions, not loading events.
  if (changeInfo.status !== 'complete') {
    return;
  }
  let matches;
  if ((matches = PRESO_REGEX.exec(tab.url)) !== null) {
    let args = matches[1];
    if (args.startsWith('?')) {
      let parts = args.substr(1);
      let params = extractParamsDictionary(parts);

      // If there is a loop parameter, then just prepare for the next page.
      if (params.loop) {
        fullscreenAndLoop(tab);
      } else {
        // if no loop parameter, highlight the pageAction button.
        chrome.pageAction.show(tab.id);
      }
    }
  }
}

chrome.pageAction.onClicked.addListener(() => {
  chrome.tabs.query({
    active: true,
    lastFocusedWindow: true
  }, function(tabs) {
    let url = tabs[0].url;
    let matches = PRESO_REGEX.exec(url);
    if (matches) {
      fullscreenAndLoop(tabs[0]);
    }
  });
});

// Listen for any changes to the URL of any tab.
chrome.tabs.onUpdated.addListener(checkForValidUrl);
