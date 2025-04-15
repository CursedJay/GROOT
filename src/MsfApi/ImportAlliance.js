const { forEach } = require('../../eslint.config.cjs');

function api_importProfileAlliance() {
  const profile = GrootApi.getAllianceProfile();
  //saveReqToFile(profile, "alliance.json")
  return profile;
}

function api_importAllianceMembers() {
  const profile = GrootApi.getAllianceMembers();
  //saveReqToFile(profile, "alliance.json")
  return profile;
}

/**
 * @return {{tcp: Number, stp: Number, totalTcp: Number}}
 */
function getAlliancePower() {
  const alliance = { members: api_importAllianceMembers(), totalTcp: 0, totalStp: 0 };
  const membersCount = alliance?.members?.length;

  if (!membersCount) return false;

  for (member of alliance.members) {
    alliance.totalTcp += member.card.tcp;
    alliance.totalStp += member.card.stp;
  }

  alliance.averageTcp = Math.floor(alliance.totalTcp / membersCount);
  alliance.averageStp = Math.floor(alliance.totalStp / membersCount);
  return { tcp: alliance.averageTcp, stp: alliance.averageStp, totalTcp: alliance.totalTcp };
}
