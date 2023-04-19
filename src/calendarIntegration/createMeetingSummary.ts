export function createMeetingSummary(groupInviteLink, organizer) {
  const wireDownloadLink = "https://wire.com/en/download/";
  const addinDownloadLink = undefined;
  const fullInvite = `<div>
    <p>${organizer} is inviting you to join this meeting in Wire.</p>
    <p>Join meeting in Wire <a href="${groupInviteLink}">${groupInviteLink}</a></p>
    <p><a href="${wireDownloadLink}">Download Wire</a></p>
    <p><a href="${addinDownloadLink}">Get Wire add-in for Outlook</a></p>
  </div>`;

  return fullInvite;
}
