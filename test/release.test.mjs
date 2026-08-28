import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import test from 'node:test';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const read = file => readFileSync(join(root, file), 'utf8');

test('ExcelAI 品牌与 v40 静态资源保持一致', () => {
  const html = read('taskpane.html');
  assert.match(html, /<title>ExcelAI<\/title>/);
  assert.match(html, /ExcelAI 设置/);
  assert.doesNotMatch(html, /DeepSeek Excel (?:助手|Assistant)/);
  for (const asset of ['app.css', 'app.js', 'icon32.png']) {
    assert.match(html, new RegExp(asset.replace('.', '\\.') + '\\?v=40'));
  }
});

test('设置中心包含五个一一对应的分类页', () => {
  const html = read('taskpane.html');
  const tabs = [...html.matchAll(/data-settings-tab="([^"]+)"/g)].map(match => match[1]);
  const panels = [...html.matchAll(/data-settings-panel="([^"]+)"/g)].map(match => match[1]);
  const expected = ['api', 'skills', 'quick', 'appearance', 'data'];
  assert.deepEqual(tabs, expected);
  assert.deepEqual(panels, expected);
  for (const id of ['provider', 'refreshModels', 'apikey', 'skillList', 'quickList', 'accent', 'storageInfo']) {
    assert.equal((html.match(new RegExp('id="' + id + '"', 'g')) || []).length, 1, id + ' 应唯一');
  }
});

test('云端与本机 manifest 使用独立 ID 和同一 v40 页面', () => {
  const cloud = read('manifest.xml');
  const local = read('manifest-standalone.xml');
  const idOf = xml => xml.match(/<Id>([^<]+)<\/Id>/)?.[1];
  assert.notEqual(idOf(cloud), idOf(local));
  for (const [name, xml] of [['cloud', cloud], ['local', local]]) {
    assert.match(xml, /<Version>1\.0\.40\.0<\/Version>/, name);
    assert.match(xml, /taskpane\.html\?v=40/, name);
    assert.doesNotMatch(xml, /\?v=39/, name);
    assert.match(xml, /<ProviderName>ExcelAI<\/ProviderName>/, name);
  }
  assert.match(cloud, /<DisplayName DefaultValue="ExcelAI"\/>/);
  assert.match(local, /<DisplayName DefaultValue="ExcelAI（本机）"\/>/);
});

test('部署脚本同步两套 manifest 并强制公网 HTTPS', () => {
  const deploy = read('deploy.mjs');
  assert.match(deploy, /\['manifest\.xml', 'manifest-standalone\.xml'\]/);
  assert.match(deploy, /parsed\.protocol !== 'https:'/);
  assert.match(deploy, /<SupportUrl DefaultValue=/);
  assert.match(deploy, /<AppDomain>/);
});

test('本机安装器发布名已更新且保留升级兼容路径', () => {
  const setup = read('setup.iss');
  assert.match(setup, /#define MyAppName "ExcelAI"/);
  assert.match(setup, /#define MyAppVersion "0\.40"/);
  assert.match(setup, /OutputBaseFilename=ExcelAI-Standalone-Setup-v0\.40/);
  assert.match(setup, /DefaultDirName=\{localappdata\}\\DeepSeekExcelAssistant/);
});
