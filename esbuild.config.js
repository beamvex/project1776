const esbuild = require('esbuild');

const buildOptions = {
  entryPoints: ['src/index.ts'],
  bundle: true,
  outfile: 'dist/index.js',
  platform: 'node',
  target: 'node16',
  format: 'cjs',
  sourcemap: true,
  minify: false,
  external: [],
};

// Build function
const build = async () => {
  try {
    await esbuild.build(buildOptions);
    console.log('Build completed successfully!');
  } catch (error) {
    console.error('Build failed:', error);
    process.exit(1);
  }
};

// Watch function for development
const watch = async () => {
  const ctx = await esbuild.context(buildOptions);
  await ctx.watch();
  console.log('Watching for changes...');
};

// Export for use in package.json scripts
module.exports = { build, watch, buildOptions };

// Run build if this file is executed directly
if (require.main === module) {
  const args = process.argv.slice(2);
  if (args.includes('--watch')) {
    watch();
  } else {
    build();
  }
}
