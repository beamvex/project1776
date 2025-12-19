# TypeScript Project with ESLint, esbuild, and Prettier

A modern TypeScript project setup with linting, building, and formatting tools.

## 🚀 Features

- **TypeScript** - Type-safe JavaScript development
- **ESLint** - Code linting with TypeScript support
- **esbuild** - Fast bundling and building
- **Prettier** - Code formatting
- **Node.js** - Runtime environment

## 📁 Project Structure

```
your-project/
├── src/
│   └── index.ts          # Main application entry point
├── dist/                 # Build output (generated)
├── esbuild.config.js     # esbuild configuration
├── tsconfig.json         # TypeScript configuration
├── eslint.config.js      # ESLint configuration
├── .prettierrc           # Prettier configuration
├── .prettierignore       # Prettier ignore patterns
├── .gitignore            # Git ignore patterns
└── package.json          # Project dependencies and scripts
```

## 🛠️ Available Scripts

### Development
- `npm run dev` - Start development with file watching
- `npm run build:watch` - Build with file watching

### Building
- `npm run build` - Build the project for production
- `npm run clean` - Remove build artifacts

### Code Quality
- `npm run lint` - Run ESLint on source files
- `npm run lint:fix` - Run ESLint and fix issues automatically
- `npm run format` - Format code with Prettier
- `npm run format:check` - Check if code is formatted correctly
- `npm run type-check` - Run TypeScript type checking

### Running
- `npm start` - Run the built application

## 🚀 Getting Started

1. Install dependencies:
   ```bash
   npm install
   ```

2. Build the project:
   ```bash
   npm run build
   ```

3. Run the application:
   ```bash
   npm start
   ```

4. For development with auto-rebuild:
   ```bash
   npm run dev
   ```

## 🔧 Configuration

### TypeScript (`tsconfig.json`)
- Target: ES2020
- Module: ESNext
- Strict mode enabled
- Source maps enabled
- Path mapping for `@/*` to `src/*`

### ESLint (`eslint.config.js`)
- TypeScript parser and plugin
- Prettier integration
- Recommended rules enabled
- Type-aware linting

### esbuild (`esbuild.config.js`)
- Entry point: `src/index.ts`
- Output: `dist/index.js`
- Platform: Node.js
- Source maps enabled
- Watch mode support

### Prettier (`.prettierrc`)
- Single quotes
- Semicolons
- 2-space indentation
- 80 character line width
- ES5 trailing commas

## 📦 Dependencies

### Development Dependencies
- `typescript` - TypeScript compiler
- `@types/node` - Node.js type definitions
- `eslint` - JavaScript/TypeScript linter
- `@typescript-eslint/parser` - TypeScript parser for ESLint
- `@typescript-eslint/eslint-plugin` - TypeScript rules for ESLint
- `prettier` - Code formatter
- `eslint-config-prettier` - Disable ESLint rules that conflict with Prettier
- `eslint-plugin-prettier` - Run Prettier as an ESLint rule
- `esbuild` - Fast JavaScript bundler and minifier
# project1776
