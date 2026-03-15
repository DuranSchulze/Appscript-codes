const ZERO_WIDTH_SPACE = "\u200B";
const LONG_TOKEN_THRESHOLD = 12;
const CHUNK_SIZE = 8;

const SOFT_BREAK_AFTER_PATTERN = /([/_\\|])/g;
const SOFT_BREAK_AROUND_PATTERN = /([\-.,;:])/g;

const splitLongToken = (token: string) => {
  if (token.length <= LONG_TOKEN_THRESHOLD) {
    return token;
  }

  const parts: string[] = [];
  for (let index = 0; index < token.length; index += CHUNK_SIZE) {
    parts.push(token.slice(index, index + CHUNK_SIZE));
  }

  return parts.join(ZERO_WIDTH_SPACE);
};

export const formatExportText = (value: string) => {
  const normalized = String(value || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(SOFT_BREAK_AFTER_PATTERN, `$1${ZERO_WIDTH_SPACE}`)
    .replace(SOFT_BREAK_AROUND_PATTERN, `${ZERO_WIDTH_SPACE}$1${ZERO_WIDTH_SPACE}`);

  return normalized
    .split(/(\s+)/)
    .map((token) => (/^\s+$/.test(token) ? token : splitLongToken(token)))
    .join("");
};
