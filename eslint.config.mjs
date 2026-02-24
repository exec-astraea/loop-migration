import eslint from "@eslint/js";
import tseslint from "typescript-eslint";

export default tseslint.config(
  eslint.configs.recommended,
  ...tseslint.configs.recommendedTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      // -- Catch real bugs --
      "no-constant-binary-expression": "error",
      "no-self-compare": "error",
      "no-template-curly-in-string": "warn",
      "no-unmodified-loop-condition": "error",

      // TypeScript already handles these; disable noisy base rules
      "no-unused-vars": "off",
      "@typescript-eslint/no-unused-vars": [
        "warn",
        { argsIgnorePattern: "^_", varsIgnorePattern: "^_" },
      ],

      // Async hygiene
      "no-return-await": "off",
      "@typescript-eslint/return-await": ["error", "in-try-catch"],
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-misused-promises": "error",
      "@typescript-eslint/require-await": "warn",

      // Safety nets
      "@typescript-eslint/no-unnecessary-condition": "warn",
      "@typescript-eslint/prefer-nullish-coalescing": "off", // we use || intentionally to catch empty strings
      "@typescript-eslint/strict-boolean-expressions": "off", // too noisy for CLI code
      "@typescript-eslint/no-unsafe-argument": "warn",
      "@typescript-eslint/no-unsafe-assignment": "warn",
      "@typescript-eslint/no-unsafe-call": "warn",
      "@typescript-eslint/no-unsafe-member-access": "warn",
      "@typescript-eslint/no-unsafe-return": "warn",

      // Style — only the genuinely useful ones
      eqeqeq: ["error", "always"],
      "prefer-const": "warn",
      "@typescript-eslint/consistent-type-imports": [
        "warn",
        { prefer: "type-imports" },
      ],

      // Relax rules that fight this codebase's patterns
      "@typescript-eslint/no-explicit-any": "off", // we use `any` deliberately in a few spots
      "@typescript-eslint/no-non-null-assertion": "off", // controlled use after checks
      "@typescript-eslint/restrict-template-expressions": "off", // template strings everywhere
    },
  },
  {
    // Test files: relax type-aware rules that are just noise in tests
    files: ["src/__tests__/**"],
    rules: {
      "@typescript-eslint/no-unsafe-argument": "off",
      "@typescript-eslint/no-unsafe-assignment": "off",
      "@typescript-eslint/no-unsafe-call": "off",
      "@typescript-eslint/no-unsafe-member-access": "off",
      "@typescript-eslint/no-unsafe-return": "off",
      "@typescript-eslint/no-floating-promises": "off",
    },
  },
  {
    // Hierarchy module intentionally works with raw untyped Fluid JSON
    files: ["src/lib/hierarchy.mts"],
    rules: {
      "@typescript-eslint/no-unsafe-argument": "off",
      "@typescript-eslint/no-unsafe-assignment": "off",
      "@typescript-eslint/no-unsafe-call": "off",
      "@typescript-eslint/no-unsafe-member-access": "off",
      "@typescript-eslint/no-unsafe-return": "off",
    },
  },
  {
    ignores: ["node_modules", "export", "export-*", "*.js", "*.mjs"],
  },
);
