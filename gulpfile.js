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

//  File Paths
const files = {
    scssPath: 'app/scss/**/*.scss',
    jsPath: 'app/js/**/*.js'
}

// Sass task: compiles the styles.css into styles.css
function scssTask() {
    return src(files.scssPath)
        .pipe(sourcemaps.init()) // initialize source maps first 
        .pipe(sass()) // compile SCSS to CSS 
        .pipe(postcss([ autoprefixer(), cssnano() ])) // PostCss plugins
        .pipe(sourcemaps.write('.')) //write sourcemaps file in current dir
        .pipe(dest('dist/assets/css') // put final css in dist folder
    );
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
    );
}

// Cachebust
var cbString = new Date().getTime();
function cacheBustTask() {
    return src(['index.html'])
        .pipe(replace(/cb=\d+/g, 'cb=' + cbString))
        .pipe(dest('.'));
}

// Watch Task: watch SCSS and JS files for changes
// If any change, run scss  and js tasks simultaneously
function watchTask() {
    watch([files.scssPath, files.jsPath],
            parallel(scssTask, jsTask));
}

// Export the default Gulp task so it can be run
// Runs the scss and js tasks simultaneously
// then runs cacheBust, then watch task
exports.default = series(
    parallel(scssTask, jsTask),
    cacheBustTask,
    watchTask
);