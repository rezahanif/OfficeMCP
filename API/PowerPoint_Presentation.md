# Presentations.Open

## Syntax
expression.Application.Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)

## Description
Opens a presentation. WithWindow=False opens hidden.

# Presentation.SaveAs

## Syntax
expression.Presentation.SaveAs(FileName)

## Description
Saves the presentation. Default format is .pptx.

# Presentation.Slides.Add

## Syntax
expression.Presentation.Slides.Add(Index, Layout)

## Description
Inserts a slide at Index (1-based). Layout: 1=title slide, 2=title+content.

## Notes (COM)
Cross-platform alternative: pptx_add_slide(path, title, content).

# Presentation.SlideShowSettings.Run

## Syntax
expression.Presentation.SlideShowSettings.Run

## Description
Starts the slideshow.
