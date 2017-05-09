<script type="text/javascript">
function maskEdit(pattern){
  var src = event.srcElement;
  var selRange = document.selection.createRange();
  var srcRange = src.createTextRange();
  selRange.setEndPoint("StartToStart", srcRange);
  var num = selRange.text + String.fromCharCode(event.keyCode) + srcRange.text.substr(selRange.text.length);
  event.returnValue = pattern.test(num); }
</script>