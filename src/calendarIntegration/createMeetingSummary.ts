export function createMeetingSummary(groupInviteLink, organizer) {
  const wireDownloadLink = "https://wire.com/en/download/";
  const fullInvite = `<div>
    <p>${organizer} is inviting you to join this meeting in Wire.</p>
    <p>Join meeting in Wire <a href="${groupInviteLink}">${groupInviteLink}</a></p>
    <p><a href="${wireDownloadLink}">Download Wire</a></p>
  </div>`;

  return fullInvite;
}
