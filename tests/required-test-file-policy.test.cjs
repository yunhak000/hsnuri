/* eslint-disable @typescript-eslint/no-require-imports */
const assert = require("node:assert/strict");
const { readFileSync } = require("node:fs");
const { join } = require("node:path");
const test = require("node:test");

test("project policy requires a new test file for every repository change", () => {
  const agents = readFileSync(join(__dirname, "..", "AGENTS.md"), "utf8");
  const policy = agents.split(/^## /m).find((section) => section.startsWith("작업별 테스트 파일 작성"));

  assert.ok(policy, "AGENTS.md must contain the per-task test policy");
  assert.match(policy, /코드·설정·문서.*모든 작업.*새 테스트 파일을 최소 1개 추가한다/);
  assert.match(policy, /기존 테스트 수정이나 실행만으로 새 파일 작성을 대체할 수 없다/);
  assert.match(policy, /`npm test`를 실행한다/);
  assert.match(policy, /테스트 파일 경로, 검증한 항목, 실행 결과를 본문에 기록한다/);
});
