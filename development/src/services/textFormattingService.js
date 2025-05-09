class TextFormattingService {
    applyBoldColor(shapes, color) {
        shapes.forEach(shape => {
            if (shape.textFrame && shape.textFrame.textRange) {
                const textRange = shape.textFrame.textRange;
                for (let i = 0; i < textRange.length; i++) {
                    const character = textRange.char(i);
                    if (character.font.bold) {
                        character.font.color = color;
                    }
                }
            }
        });
    }

    setTextColor(shapes, color) {
        shapes.forEach(shape => {
            if (shape.textFrame && shape.textFrame.textRange) {
                shape.textFrame.textRange.font.color = color;
            }
        });
    }
}

export default TextFormattingService;