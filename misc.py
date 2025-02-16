import copy
import io


def delete_slide(prs, slide):
    # Make dictionary with necessary information
    id_dict = {slide.id: [i, slide.rId] for i, slide in enumerate(prs.slides._sldIdLst)}
    slide_id = slide.slide_id
    prs.part.drop_rel(id_dict[slide_id][1])
    del prs.slides._sldIdLst[id_dict[slide_id][0]]


def duplicate_slide(pres, index):
    template = pres.slides[index]

    # Add a new slide
    copied_slide = pres.slides.add_slide(template.slide_layout)

    # Delete the existing shapes that are part of the layout
    for shp in copied_slide.shapes:
        copied_slide.shapes.element.remove(shp.element)

    # Perform a deep copy of the shapes from the template
    for shp in template.shapes:
        if "Picture" in shp.name:
            img = io.BytesIO(shp.image.blob)
            copied_slide.shapes.add_picture(
                image_file=img,
                left=shp.left,
                top=shp.top,
                width=shp.width,
                height=shp.height,
            )
        else:
            el = shp.element
            newel = copy.deepcopy(el)
            copied_slide.shapes._spTree.insert_element_before(newel, "p:extLst")

    return copied_slide


def replace_paragraph_text_retaining_initial_formatting(paragraph, new_text):
    p = paragraph._p  # the lxml element containing the `<a:p>` paragraph element
    # remove all but the first run
    for idx, run in enumerate(paragraph.runs):
        if idx == 0:
            continue
        p.remove(run._r)
    paragraph.runs[0].text = new_text
