const config = require('./webpack.config.js');
console.log("Loading config...");
try {
    config({}, { mode: 'production' }).then(c => console.log('Config loaded successfully')).catch(e => console.error('Config failed:', e));
} catch (e) {
    console.error('Sync error:', e);
}
