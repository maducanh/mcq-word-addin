function formatMCQ() {
  Word.run(async (context) => {
    const body = context.document.body;
    body.load("text");
    await context.sync();

    let text = body.text;

    text = text.replace(
      /(A\. .+)\n(B\. .+)\n(C\. .+)\n(D\. .+)/g,
      "$1    $2\n$3    $4"
    );

    body.clear();
    body.insertText(text, Word.InsertLocation.start);

    await context.sync();
    alert("Đã căn chỉnh xong!");
  });
}
