import { PAGE_MARGINS, PAGE_SIZES, PX_TO_TWIP } from './constants.js';

export function pageContentPx(options, landscape = false) {
  const size = PAGE_SIZES[options.pageSize] ?? PAGE_SIZES.A4;
  const margin = PAGE_MARGINS[options.pageMargin] ?? PAGE_MARGINS.normal;
  return {
    width: Math.floor(((landscape ? size.height : size.width) - margin.left - margin.right) / PX_TO_TWIP),
    // 给图片所在段落留一行空间。
    height: Math.floor(((landscape ? size.width : size.height) - margin.top - margin.bottom) / PX_TO_TWIP) - 32,
  };
}

export function useLandscapeImage(resource, options) {
  return options.landscapeWide === true && resource.width / resource.height >= 2.2;
}
