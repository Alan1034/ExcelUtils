import legacy from '@vitejs/plugin-legacy'
import { fileURLToPath } from 'url'
import { defineConfig, loadEnv } from 'vite'
import { ViteImageOptimizer } from 'vite-plugin-image-optimizer';
// import obfuscatorPlugin from "vite-plugin-javascript-obfuscator";
import path from 'path';
const __filenameNew = fileURLToPath(import.meta.url)
const __dirnameNew = path.dirname(__filenameNew)
const resolve = (dir) => path.resolve(__dirnameNew, dir);
// https://vitejs.dev/config/
export default defineConfig(({ command, mode }) => {
  // 根据当前工作目录中的 `mode` 加载 .env 文件
  // 设置第三个参数为 '' 来加载所有环境变量，而不管是否有 `VITE_` 前缀。
  const env = loadEnv(mode, process.cwd(), '')
  // console.log(command)
  // console.log(env)
  // console.log(env.CURRENT_ENV)
  // console.log(env.APP_ENV)
  // console.log(env.VUE_APP_BASE_API)
  return {
    // vite环境变量配置
    define: {
      "CURRENT_ENV": JSON.stringify(env.CURRENT_ENV),
    },
    server: {
      open: true,
    },
    resolve: {
      alias: {
        '@': resolve('src'),//路径化名
      },

      extensions: ['.mjs', '.js', '.ts', '.jsx', '.tsx', '.json', '.vue'],
    },

    plugins: [
      /**
       * @description: 图片压缩插件
       * @return {*}
       */
      {
        ...ViteImageOptimizer({
          /* pass your config */
        }),
        // enforce: 'pre',
        apply: 'build',
      },
      /**
       * @description: 兼容旧版本浏览器
       * @return {*}
       */
      {
        ...legacy({
          targets: ['defaults'],
        }),
        apply: 'build',
      },
      /**
      * @description: 代码压缩加密
      * @return {*}
      */
      // obfuscatorPlugin({
      //   include: [
      //     "src/indexDB/**",
      //   ],
      //   apply: "build",
      //   // debugger: true,
      //   options: {
      //     // your javascript-obfuscator options
      //     debugProtection: true,
      //     renameGlobals: true,
      //     transformObjectKeys: true
      //     // ...  [See more options](https://github.com/javascript-obfuscator/javascript-obfuscator)
      //   },
      // }),

    ],
    css: {
      preprocessorOptions: {
        less: {
          // 支持内联 JavaScript
          javascriptEnabled: true,
        },
      },
    },
    /**
     * @description: 打包时才调用
     * @return {*}
     */
    build: {
      // https://cn.vitejs.dev/guide/build.html#library-mode
      lib: {
        // Could also be a dictionary or array of multiple entry points
        // 添加打包入口文件夹
        entry: resolve(`${__dirnameNew}/src`, 'index.ts'),
        // formats: ['es', 'cjs', 'umd', 'iife'],
        name: 'Index',
        // the proper extensions will be added
        // fileName: (format) => `index.${format}.js`,
        fileName: "index",
      },
      rollupOptions: {
        // 确保外部化处理那些你不想打包进库的依赖
        external: ['vue'],
        output: {
          // 在 UMD 构建模式下为这些外部化的依赖提供一个全局变量
          globals: {
            vue: 'Vue',
          },
        },
      },
      // minify: 'terser',
      //打包环境移除console.log，debugger
      // terserOptions: {
      //   compress: {
      //     drop_console: true,
      //     drop_debugger: true,
      //   },
      // },
      //打包文件按照类型分文件夹显示（貌似会导致性能下降）
      // rollupOptions: {
      //   output: {
      //     chunkFileNames: 'js/[name]-chunk-[hash:7].js',
      //     entryFileNames: 'js/[name]-app-[hash:7].js',
      //     assetFileNames: '[ext]/[name]-chunk-[hash:7].[ext]',
      //   },
      // },
    },
  }
})
