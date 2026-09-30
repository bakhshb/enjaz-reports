"""Prune unreferenced package parts after slide removal."""
import pathlib, posixpath
from lxml import etree
A='http://schemas.openxmlformats.org/drawingml/2006/main'
P='http://schemas.openxmlformats.org/presentationml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
def q(name):return '{'+A+'}'+name

def reachable_parts(wd):
    """Remove links to discarded slides, then keep only reachable package parts."""
    root=pathlib.Path(wd)
    ns={'p':P,'r':'http://schemas.openxmlformats.org/package/2006/relationships'}
    pres=etree.parse(str(root/'ppt/presentation.xml'))
    ids={el.get('{'+R+'}id') for el in pres.findall('.//p:sldId',ns)}
    def target(name,rel):
        source=posixpath.join(posixpath.dirname(posixpath.dirname(name)),posixpath.basename(name)[:-5])
        return posixpath.normpath(posixpath.join(posixpath.dirname(source),rel.get('Target','').split('#',1)[0])).lstrip('/')
    prels='ppt/_rels/presentation.xml.rels'
    kept={target(prels,rel) for rel in etree.parse(str(root/prels)).getroot()
          if rel.get('Id') in ids}
    for path in root.rglob('*.rels'):
        name=path.relative_to(root).as_posix(); tree=etree.parse(str(path)); rels=tree.getroot(); removed=set()
        for rel in list(rels):
            if rel.get('Type','').endswith('/slide') and target(name,rel) not in kept:
                removed.add(rel.get('Id'));rels.remove(rel)
        if removed:
            tree.write(str(path),xml_declaration=True,encoding='UTF-8',standalone=True)
            source=path.parent.parent/path.name[:-5]
            if source.is_file() and source.suffix=='.xml':
                document=etree.parse(str(source))
                for link in document.xpath('//*[@r:id]',namespaces={'r':R}):
                    if link.get('{'+R+'}id') in removed and link.tag in {q('hlinkClick'),q('hlinkMouseOver')}:
                        link.getparent().remove(link)
                document.write(str(source),xml_declaration=True,encoding='UTF-8',standalone=True)
    reachable={'[Content_Types].xml'}; pending=['']
    while pending:
        source=pending.pop()
        relname=(posixpath.join(posixpath.dirname(source),'_rels',posixpath.basename(source)+'.rels') if source else '_rels/.rels')
        if relname in reachable or not (root/relname).exists():continue
        reachable.add(relname)
        for rel in etree.parse(str(root/relname)).getroot():
            if rel.get('TargetMode')=='External':continue
            part=target(relname,rel)
            if not (root/part).is_file():raise ValueError('Missing package target: '+part)
            if part not in reachable:reachable.add(part);pending.append(part)
    content_types=root/'[Content_Types].xml'; tree=etree.parse(str(content_types))
    for child in list(tree.getroot()):
        if child.tag.endswith('Override') and child.get('PartName','').lstrip('/') not in reachable:
            tree.getroot().remove(child)
    tree.write(str(content_types),xml_declaration=True,encoding='UTF-8',standalone=True)
    return reachable
