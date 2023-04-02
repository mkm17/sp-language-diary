'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`)
build.addSuppression(`Warning - [sass] src\\webparts\\diaryLang\\components\\DiaryLang.scss: filename should end with module.sass or module.scss`);

require("./spfx-versioning")(build);
build.initialize(require('gulp'));
