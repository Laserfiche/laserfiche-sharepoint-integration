// SPFx's webpack image rule (and Vite, in the Playwright component tests)
// resolves an image import to the emitted file's URL.
declare module '*.png' {
  const url: string;
  export default url;
}
declare module '*.svg' {
  const url: string;
  export default url;
}
