/**
 * Dữ liệu sơ đồ phòng máy dùng chung cho trang "Sơ đồ phòng máy tính"
 * và trang "Đăng ký chỗ ngồi".
 *
 * Mỗi phòng gồm:
 *   id            mã phòng, phải trùng với mã trong Code.gs của Apps Script
 *   name          tên hiển thị
 *   alias         tên gọi khác (tùy chọn)
 *   location      vị trí trong trường (tùy chọn)
 *   layout        "vertical" (dãy xếp dọc) hoặc "horizontal" (dãy xếp ngang)
 *   columns/rows  danh sách số hiệu dãy theo thứ tự hiển thị trái sang phải
 *                 (vertical) hoặc trên xuống dưới (horizontal)
 *   seatsPerRow   số máy mỗi dãy, mặc định 10
 *   numberFrom    "bottom" nếu máy số 1 nằm ở cuối phòng
 *   sharedSeats   false nếu phòng không cho hai học sinh ngồi chung
 *   wall          ô sát tường trái/phải: { side, label, kind }
 *   deskRow       các ô ở cuối phòng: [{ label, column, span, kind }]
 *
 * kind nhận "board" (bảng), "door" (cửa ra vào) hoặc "desk" (bàn giáo viên).
 */
window.PHONG_MAY = {
  defaultSeatsPerRow: 10,
  pairPositions: [1, 2],
  courses: [
    "TC_TIN_13",
    "TC_TIN_14",
    "TC_TIN_15",
    "TC_MOS_13",
    "TC_MOS_14",
    "TC_MOS_15"
  ],
  rooms: [
    {
      id: "PM1",
      name: "Phòng máy 1",
      location: "Tầng trệt khu B, gần tiệm photo Nguyễn Khuyến",
      layout: "vertical",
      columns: [1, 2, 3, 4],
      seatsPerRow: 12,
      numberFrom: "bottom",
      wall: { side: "right", label: "CỬA RA VÀO", kind: "door" },
      deskRow: [{ label: "BÀN GV", column: 2, span: 2, kind: "desk" }],
      note: "Bốn dãy máy xếp dọc, đánh số từ trái sang phải. Trong mỗi dãy có 12 máy, máy số 1 nằm ở cuối phòng phía bàn giáo viên, đếm ngược lên trên đến máy số 12. Cửa ra vào ở tường bên phải."
    },
    {
      id: "PM2",
      name: "Phòng máy 2",
      location: "Tầng 2 khu B",
      layout: "vertical",
      columns: [5, 4, 3, 2, 1],
      numberFrom: "bottom",
      wall: { side: "left", label: "CỬA RA VÀO", kind: "door" },
      deskRow: [{ label: "BÀN GV", column: 3, span: 2, kind: "desk" }],
      note: "Năm dãy máy xếp dọc, đánh số từ phải sang trái: dãy 1 sát bên phải, dãy 5 sát bên trái. Trong mỗi dãy, máy số 1 nằm ở cuối phòng, đếm ngược lên trên đến máy số 10. Bàn giáo viên ở cuối phòng, giữa dãy 2 và dãy 3. Cửa ra vào ở tường bên trái."
    },
    {
      id: "PM3",
      name: "Phòng máy 3",
      location: "Tầng 2 khu B",
      layout: "vertical",
      columns: [5, 4, 3, 2, 1],
      numberFrom: "bottom",
      wall: { side: "left", label: "CỬA RA VÀO", kind: "door" },
      deskRow: [{ label: "BÀN GV", column: 3, span: 2, kind: "desk" }],
      note: "Năm dãy máy xếp dọc: dãy 1 sát bên phải, dãy 5 sát bên trái. Máy số 1 nằm ở cuối phòng, đếm ngược lên trên đến máy số 10. Bàn giáo viên ở cuối phòng, giữa dãy 2 và dãy 3, cửa ra vào ở tường bên trái."
    },
    {
      id: "PM4",
      name: "Phòng máy 4",
      location: "Tầng 2 khu B",
      layout: "vertical",
      columns: [6, 5, 4, 3, 2, 1],
      numberFrom: "bottom",
      deskRow: [
        { label: "CỬA RA VÀO", column: 1, span: 1, kind: "door" },
        { label: "BÀN GV", column: 3, span: 2, kind: "desk" }
      ],
      note: "Sáu dãy máy xếp dọc, đánh số từ phải sang trái: dãy 1 nằm sát bên phải, dãy 6 sát bên trái. Trong mỗi dãy, máy số 1 nằm ở cuối phòng, đếm ngược lên trên đến máy số 10. Bàn giáo viên ở cuối phòng giữa dãy 4 và dãy 3, cửa ra vào ở góc cuối bên trái."
    },
    {
      id: "PM5",
      name: "Phòng máy 5",
      alias: "Phòng LAB",
      location: "Tầng 1 khu B",
      layout: "horizontal",
      rows: [1, 2, 3, 4, 5, 6],
      seatsPerRow: 8,
      sharedSeats: false,
      note: "Sáu dãy máy xếp ngang, mỗi dãy chia thành hai cụm 4 máy với lối đi ở giữa. Bàn giáo viên ở góc trên bên trái, bảng ở góc trên bên phải. Phòng này máy còn tốt nên mỗi học sinh ngồi một máy, không ghép chung."
    }
  ]
};

window.PHONG_MAY.seatsPerRow = function (room) {
  return room.seatsPerRow || window.PHONG_MAY.defaultSeatsPerRow;
};

window.PHONG_MAY.rowCount = function (room) {
  return (room.layout === "vertical" ? room.columns : room.rows).length;
};

window.PHONG_MAY.totalSeats = function (room) {
  return window.PHONG_MAY.rowCount(room) * window.PHONG_MAY.seatsPerRow(room);
};

window.PHONG_MAY.byId = function (id) {
  var found = null;
  window.PHONG_MAY.rooms.forEach(function (room) {
    if (room.id === id) found = room;
  });
  return found;
};
