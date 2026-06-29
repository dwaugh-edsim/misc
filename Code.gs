/**
 * Google Apps Script Webhook Backend for 40th Reunion Playlist
 *
 * TO INSTALL:
 * 1. Open a Google Sheet.
 * 2. Click Extensions > Apps Script.
 * 3. Delete any default code and paste this script.
 * 4. Click Save (disk icon).
 * 5. Click "Deploy" > "New deployment".
 * 6. Select Type: "Web app".
 * 7. Set Description, execute as "Me", and set access to "Anyone".
 * 8. Deploy, authorize permissions, and copy the Web App URL for your frontend index.html!
 */

// Initialize sheets and populate with the default 43-song playlist if empty
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Master Playlist Sheet
  let playlistSheet = ss.getSheetByName("MasterPlaylist");
  if (!playlistSheet) {
    playlistSheet = ss.insertSheet("MasterPlaylist");
    playlistSheet.appendRow(["ID", "Title", "Artist", "Year", "Description", "Votes"]);
    playlistSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#e2e8f0");
  }
  
  // 2. Setup Suggestions Sheet
  let suggestionsSheet = ss.getSheetByName("Suggestions");
  if (!suggestionsSheet) {
    suggestionsSheet = ss.insertSheet("Suggestions");
    suggestionsSheet.appendRow(["Timestamp", "Title", "Artist", "Year", "Comment"]);
    suggestionsSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#e2e8f0");
  }

  // Populate default playlist if it only contains the header
  if (playlistSheet.getLastRow() === 1) {
    const defaultSongs = [
      ["1", "I Can't Wait", "Nu Shooz", 1986, "Signature synth bassline, super danceable, and still feels incredibly fresh.", 0],
      ["2", "Word Up!", "Cameo", 1986, "An absolute funk-rock masterpiece with massive attitude that gets everyone moving.", 0],
      ["3", "Burning Down the House", "Talking Heads", 1983, "Art-school funk that brings high energy without being a standard radio cliché.", 0],
      ["4", "Bizarre Love Triangle", "New Order", 1986, "Outstanding, emotional synth-pop dance floor filler.", 0],
      ["5", "Opportunities (Let's Make Lots of Money)", "Pet Shop Boys", 1986, "Uptempo and sarcastic synth-pop classic that serves as a fantastic party track.", 0],
      ["6", "Our House", "Madness", 1982, "Incredibly catchy, upbeat, and nostalgic ska-pop classic that always gets the crowd singing along.", 0],
      ["7", "Don't You Want Me?", "The Human League", 1981, "A classic synth-pop anthem with a dramatic duet structure and a heavy, infectious beat.", 0],
      ["8", "Tainted Love", "Soft Cell", 1981, "Minimalist yet highly energetic synth-pop masterpiece that everyone knows and sings along to.", 0],
      ["9", "Super Freak", "Rick James", 1981, "The ultimate high-energy funk-disco banger with one of the most famous basslines in music history.", 0],
      ["10", "Celebration", "Kool & The Gang", 1980, "The universal, timeless anthem for happy occasions and party dance floors.", 0],
      ["11", "1999", "Prince", 1982, "Apocalyptic party funk that defined the early 80s dance sound.", 0],
      ["12", "Push It", "Salt-N-Pepa", 1986, "A pioneering hip-hop/dance crossover with an irresistible, energetic rhythm.", 0],
      ["13", "Holiday", "Madonna", 1983, "A bright, sunny, and incredibly catchy dance-pop anthem that started Madonna's rise.", 0],
      ["14", "Wild Thing", "Tone Lōc", 1988, "Classic late-80s hip-hop track featuring the recognizable Van Halen 'Jamie's Cryin' guitar hook.", 0],
      ["15", "Pump Up the Jam", "Technotronic", 1989, "A high-octane hip-house anthem that bridges the late 80s dance club transition.", 0],
      ["16", "Blue Monday", "New Order", 1983, "The best-selling 12-inch single of all time, presenting an iconic, driving electronic beat.", 0],
      ["17", "Walk Like an Egyptian", "The Bangles", 1986, "Fun, quirky pop-rock smash hit with an infectious tempo and great sing-along capability.", 0],
      ["18", "Girls Just Want to Have Fun", "Cyndi Lauper", 1983, "The quintessential energetic 80s girl-power party anthem.", 0],
      ["19", "Our Lips Are Sealed", "The Go-Go's", 1981, "Upbeat, catchy new wave pop-rock classic that captures the early 80s vibe perfectly.", 0],
      ["20", "The Safety Dance", "Men Without Hats", 1982, "Charming, high-energy synth-pop track that's guaranteed to get people dancing.", 0],
      ["21", "Let's Dance", "David Bowie", 1983, "Bowie's huge, Nile Rodgers-produced post-disco dance masterpiece.", 0],
      ["22", "My Prerogative", "Bobby Brown", 1988, "A powerful New Jack Swing anthem with a heavy beat and defiant, energetic vocals.", 0],
      ["23", "Footloose", "Kenny Loggins", 1984, "The ultimate movie soundtrack dance anthem that gets everyone on their feet instantly.", 0],
      ["24", "Venus", "Bananarama", 1986, "A high-energy, fast-paced synth-pop cover that went straight to number one.", 0],
      ["25", "Wake Me Up Before You Go-Go", "Wham!", 1984, "An incredibly energetic, upbeat pop anthem with a vibrant retro soul vibe.", 0],
      ["26", "Gloria", "Laura Branigan", 1982, "A driving, high-tempo pop-rock epic with soaring vocals and incredible drive.", 0],
      ["27", "We Got the Beat", "The Go-Go's", 1981, "Fast-paced, energetic punk-pop anthem that is a staple of early 80s party playlists.", 0],
      ["28", "Funky Cold Medina", "Tone Lōc", 1989, "Classic party rap with a heavy rock sample and humorous lyrics.", 0],
      ["29", "Mickey", "Toni Basil", 1981, "The ultimate cheerleading-style high-energy pop chant that is impossible to ignore.", 0],
      ["30", "Rumors", "Timex Social Club", 1986, "A seminal mid-80s R&B/freestyle dance track with a very catchy rhythm.", 0],
      ["31", "Express Yourself", "Madonna", 1989, "A powerful, energetic dance-pop track with a strong message and a great brass-driven groove.", 0],
      ["32", "Hot Hot Hot", "Buster Poindexter", 1987, "A raucous, horn-filled party song that is the quintessential festive conga-line starter.", 0],
      ["33", "Dancing in the Dark", "Bruce Springsteen", 1984, "Springsteen's biggest synth-rock hit with an upbeat tempo and a classic 80s drum beat.", 0],
      ["34", "99 Red Balloons", "Nena", 1983, "The iconic, high-energy German-turned-global new wave hit that everyone knows.", 0],
      ["35", "Come On Eileen", "Dexys Midnight Runners", 1982, "Unstoppable celtic-pop-rock sing-along party classic with a legendary tempo increase.", 0],
      ["36", "You Spin Me Round (Like a Record)", "Dead or Alive", 1984, "Pure, high-octane Hi-NRG dance floor dynamite with a massive electronic drive.", 0],
      ["37", "Black Cars", "Gino Vannelli", 1985, "Sleek, synth-heavy Canadian pop-rock track with a very cool mid-80s groove.", 0],
      ["38", "Doesn't Really Matter", "Platinum Blonde", 1985, "A massive, uptempo Canadian new-wave/pop-rock hit with a very catchy synth hook and energetic beat.", 0],
      ["39", "The Sun Always Shines on T.V.", "a-ha", 1985, "A dramatic, soaring synth-pop masterpiece with a massive, high-energy beat.", 0],
      ["40", "Beat It", "Michael Jackson", 1982, "The legendary high-energy dance-rock crossover with Eddie Van Halen's iconic guitar solo.", 0],
      ["41", "We Don't Have to Take Our Clothes Off", "Jermaine Stewart", 1986, "An incredibly infectious, upbeat dance-pop anthem that is a quintessential 80s party tune.", 0],
      ["42", "Let's Talk About Sex", "Salt-N-Pepa", 1990, "Playful, catchy hip-hop track with a massive groove that gets the room talking and moving.", 0],
      ["43", "Faith", "George Michael", 1987, "An iconic, bopping pop-rock classic with that unmistakable organ intro and acoustic guitar groove.", 0]
    ];
    playlistSheet.getRange(2, 1, defaultSongs.length, 6).setValues(defaultSongs);
  }
}

// Helper to construct JSON responses with CORS headers
function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// GET Endpoint: Return current playlist and votes
function doGet(e) {
  try {
    setupDatabase(); // Ensure DB is initialized
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("MasterPlaylist");
    const data = sheet.getDataRange().getValues();
    
    const playlist = [];
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      playlist.push({
        id: String(data[i][0]),
        title: String(data[i][1]),
        artist: String(data[i][2]),
        year: Number(data[i][3]),
        description: String(data[i][4]),
        votes: Number(data[i][5] || 0)
      });
    }
    
    return jsonResponse({ status: "success", playlist: playlist });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// POST Endpoint: Handle Upvotes, Downvotes, and Song Suggestions
function doPost(e) {
  try {
    setupDatabase();
    
    // Parse request body
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (action === "vote") {
      const songId = String(postData.id);
      const voteType = postData.type; // "up" or "down"
      const delta = voteType === "up" ? 1 : -1;
      
      const sheet = ss.getSheetByName("MasterPlaylist");
      const data = sheet.getDataRange().getValues();
      let found = false;
      let newVotes = 0;
      
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === songId) {
          const voteCell = sheet.getRange(i + 1, 6);
          newVotes = Number(data[i][5] || 0) + delta;
          voteCell.setValue(newVotes);
          found = true;
          break;
        }
      }
      
      if (!found) {
        return jsonResponse({ status: "error", message: "Song ID not found." });
      }
      
      return jsonResponse({ status: "success", id: songId, votes: newVotes });
      
    } else if (action === "suggest") {
      const title = postData.title;
      const artist = postData.artist;
      const year = postData.year;
      const comment = postData.comment;
      
      if (!title || !artist) {
        return jsonResponse({ status: "error", message: "Title and Artist are required." });
      }
      
      const sheet = ss.getSheetByName("Suggestions");
      sheet.appendRow([new Date(), title, artist, year, comment]);
      
      return jsonResponse({ status: "success", message: "Suggestion added successfully!" });
    }
    
    return jsonResponse({ status: "error", message: "Invalid action." });
    
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}
