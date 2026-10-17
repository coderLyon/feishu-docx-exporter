/** 截当前视口，便于判断菜单/对话框状态 */
import fs from 'node:fs/promises';
import { Cdp, PORT } from './cdp.mjs';

const cdp = await Cdp.connect(PORT);
const list = await cdp.targets();
const page = list.find((t) => t.type === 'page');
const { sessionId } = await cdp.send('Target.attachToTarget', { targetId: page.id, flatten: true }, null);
cdp.sessionId = sessionId;
await cdp.send('Page.enable');
const shot = await cdp.send('Page.captureScreenshot', { format: 'png' });
const out = 'samples/out/state.png';
await fs.writeFile(out, Buffer.from(shot.data, 'base64'));
console.log('已保存', out);
cdp.close();
