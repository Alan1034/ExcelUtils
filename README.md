# VueNpmFrame

一个从零搭建的Vue3纯净框架，用于Npm包 <br/>
A Vue3 pure frame build up from nothing and use for npm package

以package.json内的name是"vue-npm-frame"为例,首先在根目录下输入 yarn link。然后切换到用户项目根目录下输入yarn link "vue-npm-frame"，然后启动项目。<br/>
take the name like "vue-npm-frame" in the package.json file as an example, first enter yarn link in the root directory. Then switch to the root directory of the user project, enter yarn link "vue-npm-frame" and start the project.


如果项目出现警告[Vue warn]: Invalid VNode type: Symbol("Text") (symbol) ，则需要在用户项目webpack中配置：<br/>
if warning like [Vue warn]: Invalid VNode type: Symbol("Text") (symbol) appears in the project, it needs to be configured in the user project webpack:

    resolve: {
        symlinks: false, //在yarn link 入其他模块的时候统一使用同一个vue引用，避免vue3的一个报错
        alias: {
            'vue$': 'vue/dist/vue.esm-bundler.js',
            vue: path.resolve(__dirname, `../node_modules/vue`) // 定义vue路径
        },
    },

取消链接：yarn unlink<br/>
cancel link: yarn unlink 

发布：npm publish<br/>
publish: npm publish

迭代： npm version [patch,minor,major]，然后 npm publish<br/>
patch： 修复bug、微小改动，改变版本号第三位<br/>
minor： 上线新功能，并对当前版本已有功能模块不影响，改变版本号第二位<br/>
major： 上线多个新功能模块，并对当前版本已有功能有影响，改变版本号第一位<br/>
iteration: npm version [patch,minor,major], then npm publish<br/>
patch: fix bugs, make little changes, and change the third digit of the version number. <br/>
major: new functions will be launched, and the existing function modules of the current version will not be affected. The second digit of the version number will be changed.<br/>
major: several new function modules will be launched, which will affect the existing functions of current version. The first digit of the version number will be changed.


安装：npm i vue-npm-frame<br/>
install: npm i vue-npm-frame