// The typed `*.module.scss.ts` files that TypeScript normally resolves these
// imports to are generated (and gitignored) by the gulp build, which the
// component-test job never runs. This fallback gives `typecheck:ct` the shape
// that stubs/scss-module.ts provides at runtime; when the generated file does
// exist, TypeScript resolves to it instead and checks the real class names.
declare module '*.module.scss' {
  const styles: Record<string, string>;
  export default styles;
}
