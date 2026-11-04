import { describe, expect, it } from 'vitest'
import { patchMasterTextStylesXml } from '../src/master-text-styles'

const MASTER = (tx: string) =>
  `<p:sldMaster xmlns:a="a" xmlns:p="p"><p:cSld/><p:txStyles>${tx}<p:otherStyle/></p:txStyles></p:sldMaster>`

describe('patchMasterTextStylesXml', () => {
  it('sets sz and b on an existing defRPr, keeping its other attributes and children', () => {
    const xml = MASTER(
      '<p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/></a:defRPr></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:defRPr sz="2800"/></a:lvl1pPr></p:bodyStyle>',
    )
    const out = patchMasterTextStylesXml(xml, {
      title: { sizePt: 40, bold: true },
      body: { 1: { sizePt: 20 } },
    })
    const titleRPr = /<p:titleStyle><a:lvl1pPr algn="l">(<a:defRPr[^>]*>)/.exec(out)![1]!
    expect(titleRPr).toContain('sz="4000"')
    expect(titleRPr).toContain('b="1"')
    expect(titleRPr).toContain('kern="1200"')
    expect(out).toContain('<a:latin typeface="+mj-lt"/>')
    expect(out).toContain('<p:bodyStyle><a:lvl1pPr><a:defRPr sz="2000"/></a:lvl1pPr></p:bodyStyle>')
  })

  it('creates missing levels in order and defRPr before extLst', () => {
    const xml = MASTER(
      '<p:titleStyle/><p:bodyStyle><a:lvl1pPr marL="0"><a:buNone/><a:extLst/></a:lvl1pPr><a:lvl3pPr/></p:bodyStyle>',
    )
    const out = patchMasterTextStylesXml(xml, {
      title: { sizePt: 36 },
      body: { 1: { bold: false }, 2: { sizePt: 18 } },
    })
    expect(out).toContain(
      '<p:titleStyle><a:lvl1pPr><a:defRPr sz="3600"/></a:lvl1pPr></p:titleStyle>',
    )
    expect(out).toContain('<a:lvl1pPr marL="0"><a:buNone/><a:defRPr b="0"/><a:extLst/></a:lvl1pPr>')
    // lvl2 lands between lvl1 and lvl3
    expect(out).toMatch(/<\/a:lvl1pPr><a:lvl2pPr><a:defRPr sz="1800"\/><\/a:lvl2pPr><a:lvl3pPr\/>/)
  })

  it('leaves the XML untouched for an empty spec', () => {
    const xml = MASTER('<p:titleStyle><a:lvl1pPr/></p:titleStyle>')
    expect(patchMasterTextStylesXml(xml, {})).toBe(xml)
  })
})
