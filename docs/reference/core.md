# `pptx_plus.core`

The foundation layer. Everything that touches the OPC package directly lives
here, and it depends on nothing above it.

!!! warning "Not the public API"
    `pptx_plus.slides` is what the compatibility promise covers. `core` is
    documented because the reasoning in it is the substance of the library,
    and because extending `pptx_plus` means working here.

## The copy engine

::: pptx_plus.core.clone.clone_part_graph

::: pptx_plus.core.clone.ClonePolicy

::: pptx_plus.core.clone.CloneResult

## Relationship-id rewriting

::: pptx_plus.core.relmap.remap_rel_ids

::: pptx_plus.core.relmap.DanglingRelationshipError

::: pptx_plus.core.relmap.RelIdLiteralWarning

## Classifying the part graph

::: pptx_plus.core.partgraph.classify

::: pptx_plus.core.partgraph.rel_edges

::: pptx_plus.core.partgraph.Disposition

::: pptx_plus.core.partgraph.RelEdge

## Parts and part names

::: pptx_plus.core.parts.clone_part

::: pptx_plus.core.parts.allocate_partname

::: pptx_plus.core.parts.drop_relationship

## Sections and custom shows

::: pptx_plus.core.sections.scrub_slide

::: pptx_plus.core.sections.reorder_slide

::: pptx_plus.core.sections.insert_slide

::: pptx_plus.core.sections.section_lst

::: pptx_plus.core.sections.custom_show_lst

## Slide ids

::: pptx_plus.core.ids.next_slide_id

::: pptx_plus.core.ids.validate_slide_id

::: pptx_plus.core.ids.SlideIdRangeError

## Elements and namespaces

::: pptx_plus.core.oxml.el

::: pptx_plus.core.oxml.xpath

::: pptx_plus.core.ns.qn

## Errors

::: pptx_plus.core.errors.PptxPlusError

::: pptx_plus.core._compat.UpstreamSurfaceError
