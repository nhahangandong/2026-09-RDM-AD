/**
 * Function Helper làm sạch chuỗi
 */
function cleanCodeValue_(val) {
  if (val === null || val === undefined) return "";
  return val.toString().trim();
}
