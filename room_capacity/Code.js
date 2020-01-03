function ROOM_TYPE(numParticipants) {
    if (numParticipants <= 25) {
      return '🎈 Medium room: 25max';
    }
    if (numParticipants <= 40) {
      return '🎉 Large room: 40max';
    }
    return '💥 Extra-large room: 200max'
  }
  