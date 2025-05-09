class ShapeResizingService {
    matchSize(shapes) {
        if (shapes.length < 2) return;

        const referenceShape = shapes[0];
        const referenceWidth = referenceShape.width;
        const referenceHeight = referenceShape.height;

        shapes.forEach(shape => {
            shape.width = referenceWidth;
            shape.height = referenceHeight;
        });
    }

    matchWidth(shapes) {
        if (shapes.length < 2) return;

        const referenceShape = shapes[0];
        const referenceWidth = referenceShape.width;

        shapes.forEach(shape => {
            shape.width = referenceWidth;
        });
    }

    matchHeight(shapes) {
        if (shapes.length < 2) return;

        const referenceShape = shapes[0];
        const referenceHeight = referenceShape.height;

        shapes.forEach(shape => {
            shape.height = referenceHeight;
        });
    }
}

export default ShapeResizingService;