/**
 * Assigns a type of room needed for a session based on its number of
 * registered participants.
 *
 * @param {number} numParticipants - number of participants.
 * 
 * @returns {string} - size of room.
 */
function ROOM_TYPE(numParticipants) {
  if (numParticipants <= 25) {
    return '🎈 Medium room: 25max';
  }
  if (numParticipants <= 40) {
    return '🎉 Large room: 40max';
  }
  return '💥 Extra-large room: 200max'
}
