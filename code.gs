/**
 * ============================================================================
 * Project: Sornorpoom Learning Portal (YouTube Web App)
 * Author: ศน.ภูมิ (สำนักงานศึกษาธิการจังหวัดเชียงใหม่)
 * File: code.gs
 * Description: สคริปต์ฝั่ง Server สำหรับ Sync ข้อมูลจาก YouTube และประมวลผลส่ง Web App
 * ============================================================================
 */

// Channel ID ของช่อง sornorpoom
const CHANNEL_ID = 'UCxCBjGfJL25wnIgsOG8KUCw';

/**
 * 1. ฟังก์ชันแสดงผลหน้า Web App ( doGET )
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sornorpoom Learning Portal')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. ฟังก์ชัน Sync ข้อมูลจาก YouTube API ลง Google Sheets
 * (ควรตั้งเวลา Time-driven Trigger ให้รันวันละ 1 ครั้ง)
 */
function syncYouTubeDataCustom() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playlistSheet = ss.getSheetByName('PLAYLISTS');
  const videoSheet = ss.getSheetByName('VIDEOS');
  
  if (!playlistSheet || !videoSheet) {
    Logger.log("⚠️ กรุณาตรวจสอบว่ามี Sheet ชื่อ 'PLAYLISTS' และ 'VIDEOS'");
    return;
  }

  // ดึง Playlists ทั้งหมด
  let playlists = [];
  let nextPageToken = '';
  
  do {
    let response = YouTube.Playlists.list('snippet,status', {
      channelId: CHANNEL_ID,
      maxResults: 50,
      pageToken: nextPageToken
    });
    if (response.items) playlists = playlists.concat(response.items);
    nextPageToken = response.nextPageToken;
  } while (nextPageToken);

  let playlistRows = [];
  let videoRows = [];
  let existingVideoKeys = getExistingVideoKeys(videoSheet);

  // วนลูปอ่านข้อมูล Playlists และ Videos
  playlists.forEach(pl => {
    let plId = pl.id;
    let plTitle = pl.snippet.title;
    let plDesc = (pl.snippet.description || "").replace(/\n/g, " ");
    let plThumb = pl.snippet.thumbnails && pl.snippet.thumbnails.high ? pl.snippet.thumbnails.high.url : '';
    
    // โครงสร้าง A-G ของ PLAYLISTS
    playlistRows.push([
      plId,        // A: Playlist_ID
      plTitle,     // B: Playlist_Title
      plDesc,      // C: Description
      plThumb,     // D: Thumbnail_URL
      '',          // E: Category_ID (ระบุหมวดหมู่หลักใน Sheet)
      '',          // F: Target_Audience
      'Active'     // G: Status
    ]);

    // ดึงวิดีโอใน Playlist นี้
    let vNextToken = '';
    do {
      let vResponse = YouTube.PlaylistItems.list('snippet', {
        playlistId: plId,
        maxResults: 50,
        pageToken: vNextToken
      });

      if (vResponse.items) {
        vResponse.items.forEach(item => {
          let vId = item.snippet.resourceId.videoId;
          let uniqueKey = vId + '_' + plId;
          
          if (!existingVideoKeys.has(uniqueKey)) {
            let vTitle = item.snippet.title;
            let vDesc = item.snippet.description || "";
            let publishedAt = item.snippet.publishedAt;
            let vThumb = item.snippet.thumbnails && item.snippet.thumbnails.high ? item.snippet.thumbnails.high.url : '';
            
            // สกัดดึงลิงก์ Google Drive หรือเอกสารจาก Description อัตโนมัติ (ถ้ามี)
            let driveUrlMatch = vDesc.match(/https:\/\/(drive|docs)\.google\.com\/[^\s]+/g);
            let resourceLink = driveUrlMatch ? driveUrlMatch[0] : '';

            // โครงสร้าง A-I ของ VIDEOS
            videoRows.push([
              vId,          // A: Video_ID
              plId,         // B: Playlist_ID
              vTitle,       // C: Video_Title
              vDesc,        // D: Description
              publishedAt,  // E: Published_At
              '',           // F: Duration
              vThumb,       // G: Thumbnail_URL
              resourceLink, // H: Resource_Link
              'Active'      // I: Status
            ]);
            existingVideoKeys.add(uniqueKey);
          }
        });
      }
      vNextToken = vResponse.nextPageToken;
    } while (vNextToken);
  });

  // บันทึกข้อมูลกลับลง Sheet PLAYLISTS (เขียนทับตั้งแต่แถว 2)
  if (playlistRows.length > 0) {
    playlistSheet.getRange(2, 1, playlistSheet.getLastRow() > 1 ? playlistSheet.getLastRow() - 1 : 1, 7).clearContent();
    playlistSheet.getRange(2, 1, playlistRows.length, 7).setValues(playlistRows);
  }

  // บันทึกข้อมูลเพิ่มลง Sheet VIDEOS (ต่อท้ายรายการเดิม)
  if (videoRows.length > 0) {
    let lastRow = videoSheet.getLastRow();
    videoSheet.getRange(lastRow + 1, 1, videoRows.length, 9).setValues(videoRows);
  }

  Logger.log(`✅ อัปเดตข้อมูลเรียบร้อย! Sync Playlist ${playlistRows.length} รายการ และเพิ่ม Video ใหม่ ${videoRows.length} รายการ`);
}

/**
 * 3. ฟังก์ชันดึงข้อมูลส่งให้ Web App (โหลดเร็ว: ดึงเฉพาะคลิปย้อนหลังไม่เกิน 2 ปี + Status = Active)
 */
function getWebData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const playlistSheet = ss.getSheetByName('PLAYLISTS');
  const videoSheet = ss.getSheetByName('VIDEOS');

  let playlists = [];
  let videos = [];
  let categoriesSet = new Set();
  let activePlaylistIds = new Set();

  // กำหนดขอบเขตเวลาย้อนหลัง 2 ปี นับจากวันที่ปัจจุบัน
  const now = new Date();
  const twoYearsAgo = new Date();
  twoYearsAgo.setFullYear(now.getFullYear() - 2);
  const twoYearsAgoTime = twoYearsAgo.getTime();

  // อ่านข้อมูล PLAYLISTS
  if (playlistSheet && playlistSheet.getLastRow() > 1) {
    let pData = playlistSheet.getRange(2, 1, playlistSheet.getLastRow() - 1, 7).getValues();
    pData.forEach(r => {
      let plId = r[0];
      let categoryName = (r[4] || 'อื่นๆ').toString().trim();
      let status = (r[6] || 'Active').toString().trim();

      if (status.toLowerCase() === 'active') {
        activePlaylistIds.add(plId);
        if (categoryName) categoriesSet.add(categoryName);

        playlists.push({
          id: plId,
          title: r[1],
          description: r[2],
          thumbnail: r[3],
          category: categoryName,
          target: r[5],
          status: status
        });
      }
    });
  }

  // อ่านข้อมูล VIDEOS
  if (videoSheet && videoSheet.getLastRow() > 1) {
    let lastCol = Math.max(videoSheet.getLastColumn(), 9);
    let vData = videoSheet.getRange(2, 1, videoSheet.getLastRow() - 1, lastCol).getValues();

    vData.forEach(r => {
      let vPlaylistId = r[1];
      let dateObj = r[4] ? new Date(r[4]) : new Date(0);
      let dateTime = dateObj.getTime();
      let vStatus = (r[8] || 'Active').toString().trim(); // คอลัมน์ I

      // เงื่อนไข: คลิปย้อนหลังไม่เกิน 2 ปี + Status ไม่ใช่ Hide + Playlist ต้นทางต้อง Active
      if (dateTime >= twoYearsAgoTime && vStatus.toLowerCase() !== 'hide' && activePlaylistIds.has(vPlaylistId)) {
        videos.push({
          id: r[0],
          playlistId: vPlaylistId,
          title: r[2],
          description: r[3],
          rawDate: dateTime,
          publishedAt: r[4] ? dateObj.toLocaleDateString('th-TH', { year: 'numeric', month: 'short', day: 'numeric' }) : '',
          duration: r[5],
          thumbnail: r[6],
          resourceLink: r[7]
        });
      }
    });

    // เรียงลำดับคลิปจาก ใหม่สุด -> เก่าสุด
    videos.sort((a, b) => b.rawDate - a.rawDate);
  }

  return { 
    categories: Array.from(categoriesSet), 
    playlists: playlists, 
    videos: videos 
  };
}

/**
 * ฟังก์ชันช่วยตรวจสอบ Key ซ้ำ
 */
function getExistingVideoKeys(sheet) {
  let lastRow = sheet.getLastRow();
  let set = new Set();
  if (lastRow > 1) {
    let data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    data.forEach(row => set.add(row[0] + '_' + row[1]));
  }
  return set;
}
