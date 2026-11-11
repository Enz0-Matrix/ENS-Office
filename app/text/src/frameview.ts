// Вид надписи в редакторе: ProseMirror не должен считать «правкой» то, что floats.ts ставит рамке после вёрстки
// (положение left/top, поля обтекания) — иначе он перерисовывает рамку и всё сбрасывает.

import type { Node as PMNode } from 'prosemirror-model';
import type { NodeView, ViewMutationRecord } from 'prosemirror-view';
import { schema } from './schema';

const KEEP = ['left', 'top', 'float', 'marginLeft', 'marginRight', 'marginTop'] as const;

export class FrameView implements NodeView {
  dom: HTMLElement;
  contentDOM: HTMLElement;
  private node: PMNode;

  constructor(node: PMNode) {
    this.node = node;
    this.dom = document.createElement('div');
    this.contentDOM = document.createElement('div');
    this.dom.append(this.contentDOM);
    this.apply(false);
  }

  private apply(keep: boolean): void {
    const spec = schema.nodes.frame.spec.toDOM!(this.node) as unknown as [string, Record<string, string>, [string, Record<string, string>, number]];
    const kept: Record<string, string> = {};
    if (keep) for (const k of KEEP) kept[k] = this.contentDOM.style[k];
    this.dom.className = spec[1].class;
    const attrs = spec[2][1];
    for (const n of this.contentDOM.getAttributeNames()) if (!(n in attrs) && n !== 'contenteditable') this.contentDOM.removeAttribute(n);
    for (const [k, v] of Object.entries(attrs)) this.contentDOM.setAttribute(k, v);
    if (keep) for (const k of KEEP) if (kept[k]) this.contentDOM.style[k] = kept[k];
  }

  update(node: PMNode): boolean {
    if (node.type !== this.node.type) return false;
    this.node = node;
    this.apply(true);
    return true;
  }

  ignoreMutation(m: ViewMutationRecord): boolean {
    return m.type === 'attributes';
  }
}

export const frameViews = { frame: (node: PMNode) => new FrameView(node) };
