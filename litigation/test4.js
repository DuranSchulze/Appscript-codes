let word = "Compliance [bunao]";
word = word.split(" ").filter(Boolean).map(word =>
  word.replace(/[A-Za-z][A-Za-z'`-]*/g, (segment) => {
    if (/^(?:[ivxlcdm]+)$/i.test(segment)) {
      return segment.toUpperCase();
    }
    return segment.charAt(0).toUpperCase() + segment.slice(1).toLowerCase();
  })
).join(" ");
console.log(word);
