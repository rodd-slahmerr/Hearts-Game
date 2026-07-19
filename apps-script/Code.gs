// This file is a reference copy of the Google Apps Script Web App code that
// powers the "cloud database" behind index.html (WEB_APP_URL). The live,
// authoritative version lives inside the Google Sheet's own Apps Script
// project (Extensions > Apps Script) — this copy exists so the logic is
// visible and version-controlled alongside the rest of the app. If you
// change one, update the other.
//
// Expected sheets in the spreadsheet:
//   "Players" — columns: [id, name, status ("Active"/other)]
//   "Scores"  — columns: [Game_ID, Round_Number, Timestamp, PlayerID,
//                          Tournament_Play, Cancellation_Rules, Trick_Score]

function doGet(e) {
  const action = e.parameter.action;
  if (action === "getStats") {
    return getStats();
  }
  return getPlayers();
}

function getPlayers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName("Players");
  const data = playerSheet.getDataRange().getValues();
  const players = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === "Active") {
      players.push({
        id: data[i][0],
        name: data[i][1]
      });
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      status: "success",
      players: players
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playerSheet = ss.getSheetByName("Players");
  const scoreSheet = ss.getSheetByName("Scores");

  const playerData = playerSheet.getDataRange().getValues();
  const playerNames = {};
  for (let i = 1; i < playerData.length; i++) {
    playerNames[playerData[i][0]] = playerData[i][1];
  }

  const scoreData = scoreSheet.getDataRange().getValues();

  // Group rows by Game_ID, summing each player's Trick_Score
  const games = {}; // Game_ID -> { totals: {playerId: sum}, lastTimestamp }
  for (let i = 1; i < scoreData.length; i++) {
    const row = scoreData[i];
    const gameId = row[0];
    const timestamp = row[2];
    const playerId = row[3];
    const trickScore = Number(row[6]) || 0;

    if (!gameId || !playerId) continue;

    if (!games[gameId]) {
      games[gameId] = { totals: {}, lastTimestamp: timestamp };
    }
    games[gameId].totals[playerId] = (games[gameId].totals[playerId] || 0) + trickScore;
    if (new Date(timestamp) > new Date(games[gameId].lastTimestamp)) {
      games[gameId].lastTimestamp = timestamp;
    }
  }

  const playerStats = {}; // playerId -> aggregate stats
  function ensurePlayer(id) {
    if (!playerStats[id]) {
      playerStats[id] = {
        id: id,
        name: playerNames[id] || id,
        gamesPlayed: 0,
        wins: 0,
        bestScore: null,
        worstScore: null
      };
    }
    return playerStats[id];
  }

  let bestGame = null; // lowest winning score ever recorded
  const gameSummaries = [];

  Object.keys(games).forEach(function (gameId) {
    const game = games[gameId];
    const entries = Object.keys(game.totals).map(function (pid) {
      return { playerId: pid, total: game.totals[pid] };
    });
    if (entries.length === 0) return;

    entries.sort(function (a, b) { return a.total - b.total; }); // lowest score wins
    const winnerEntry = entries[0];

    entries.forEach(function (entry) {
      const stats = ensurePlayer(entry.playerId);
      stats.gamesPlayed++;
      if (stats.bestScore === null || entry.total < stats.bestScore) stats.bestScore = entry.total;
      if (stats.worstScore === null || entry.total > stats.worstScore) stats.worstScore = entry.total;
    });

    ensurePlayer(winnerEntry.playerId).wins++;

    if (!bestGame || winnerEntry.total < bestGame.score) {
      bestGame = {
        playerId: winnerEntry.playerId,
        playerName: playerNames[winnerEntry.playerId] || winnerEntry.playerId,
        score: winnerEntry.total,
        gameId: gameId,
        timestamp: game.lastTimestamp
      };
    }

    gameSummaries.push({
      gameId: gameId,
      timestamp: game.lastTimestamp,
      winnerName: playerNames[winnerEntry.playerId] || winnerEntry.playerId,
      winnerScore: winnerEntry.total,
      players: entries.map(function (e) {
        return { name: playerNames[e.playerId] || e.playerId, score: e.total };
      })
    });
  });

  const leaderboard = Object.keys(playerStats).map(function (id) {
    const s = playerStats[id];
    s.winRate = s.gamesPlayed > 0 ? s.wins / s.gamesPlayed : 0;
    return s;
  });

  leaderboard.sort(function (a, b) {
    return b.wins - a.wins || b.winRate - a.winRate;
  });

  gameSummaries.sort(function (a, b) {
    return new Date(b.timestamp) - new Date(a.timestamp);
  });

  return ContentService
    .createTextOutput(JSON.stringify({
      status: "success",
      totalGames: Object.keys(games).length,
      leaderboard: leaderboard,
      bestGameEver: bestGame,
      recentGames: gameSummaries.slice(0, 10)
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents || "{}");
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: "Invalid JSON" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (data.action === "addPlayer") {
    const playerSheet = ss.getSheetByName("Players");
    playerSheet.appendRow([data.id, data.name, "Active"]);
    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (data.rows && Array.isArray(data.rows)) {
    const scoreSheet = ss.getSheetByName("Scores");
    data.rows.forEach(function(row) {
      scoreSheet.appendRow([
        row.Game_ID,
        row.Round_Number,
        row.Timestamp,
        row.PlayerID,
        row.Tournament_Play,
        row.Cancellation_Rules,
        row.Trick_Score
      ]);
    });
    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ status: "error", message: "Invalid request" }))
    .setMimeType(ContentService.MimeType.JSON);
}
