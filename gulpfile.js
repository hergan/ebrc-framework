// Initialize modules
// Importing specific gulp APL funcs lets us write them below as a series() instead of gulp.series()
const { src, dest, watch, series, parallel } = require('gulp');
// Importing all the gulp-related packages we want to use
const sourcemaps = require('gulp-sourcemaps');
const sass = require('gulp-sass');
const concat = require('gulp-concat');
const uglify = require('gulp-uglify');
const postcss = require('gulp-postcss');
const autoprefixer = require('autoprefixer');
const cssnano = require('cssnano');
const replace = require('gulp-replace');
const browserSync = require('browser-sync').create();

//  File Paths
const files = {
    scssPath: 'app/scss/**/*.scss',
    // css: 'dist/css',
    jsPath: 'app/js/**/*.js'
}

// A simple task to reload the page
function reload() {
    browserSync.reload();
}

// Init a bs server
function init(){
    browserSync.init({
        // You can tell browserSync to use this directory and serve it as a mini-server
        server: {
            baseDir: "./"
        }
        // If you are already serving your website locally using something like apache
        // You can use the proxy setting to proxy that instead
        // proxy: "yourlocal.dev"
    })
}

// Sass task: compiles the styles.css into styles.css
function scssTask() {
    return src(files.scssPath)
        .pipe(sourcemaps.init()) // initialize source maps first 
        .pipe(sass()) // compile SCSS to CSS 
        .on("error", sass.logError)
        .pipe(postcss([ autoprefixer(), cssnano() ])) // PostCss plugins
        .pipe(sourcemaps.write('.')) //write sourcemaps file in current dir
        .pipe(dest('dist/assets/css')) // put final css in dist folder
        .pipe(browserSync.stream());
    }

// JS Task: concatenates and uglifies JS files to script.js
function jsTask() {
    return src([
        files.jsPath
        //,'!' + 'includes/js/jquery.min.js', // to exclude any specific files
        ])
        .pipe(concat('all.js'))
        .pipe(uglify())
        .pipe(dest('dist/assets/js')
        .pipe(browserSync.stream())
    );
}

// Cachebust
var cbString = new Date().getTime();
function cacheBustTask() {
    return src(['index.html'])
        .pipe(replace(/cb=\d+/g, 'cb=' + cbString))
        .pipe(dest('.'));
}

//
function indexTask() {
    watch("./*.html", reload);
}

// Watch Task: watch SCSS and JS files for changes
// If any change, run scss  and js tasks simultaneously
function watchTask() {
    watch([files.scssPath, files.jsPath],
        parallel(scssTask, jsTask, reload));
}



// Export the default Gulp task so it can be run
// Runs the scss and js tasks simultaneously
// then runs cacheBust, then watch task
exports.default = series(
    parallel(scssTask, jsTask),
    cacheBustTask,
    parallel(init, watchTask, indexTask)
);