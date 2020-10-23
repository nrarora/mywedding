function WriteToFile(passForm) {
  var fso = CreateObject("Scripting.FileSystemObject");
  var s = fso.CreateTextFile("../RSVP.txt", True);
  s.writeline(document.passForm.input1.value);
  s.writeline(document.passForm.input2.value);
  s.writeline(document.passForm.input3.value);
  s.Close();
}
