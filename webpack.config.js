module.exports = {
  entry: './index.js',
  node: {
    global: true
  },
  externals: {
    fs: require('fs')
  },
  output: {
    filename: 'xlsx-stream-reader.bundle.js'       
  }
};
