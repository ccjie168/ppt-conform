"""
OverflowAdjuster: detect and adjust objects that overflow page boundaries.
"""

from pptx.enum.shapes import MSO_SHAPE_TYPE


class OverflowAdjuster:
    """Detect and adjust overflowing objects."""

    def __init__(self, slide_width, slide_height):
        self.slide_width = slide_width
        self.slide_height = slide_height

    def check_overflow(self, shape) -> tuple[bool, str]:
        """Check if shape overflows and return (is_overflow, direction)."""
        right = shape.left + shape.width
        bottom = shape.top + shape.height

        overflow_right = right > self.slide_width
        overflow_bottom = bottom > self.slide_height
        overflow_left = shape.left < 0
        overflow_top = shape.top < 0

        if overflow_right and overflow_bottom:
            return True, "bottom-right"
        elif overflow_right:
            return True, "right"
        elif overflow_bottom:
            return True, "bottom"
        elif overflow_left:
            return True, "left"
        elif overflow_top:
            return True, "top"
        return False, "none"

    def adjust_shape(self, shape, max_scale_down=0.7) -> bool:
        """Adjust overflowing shape: scale down or move to fit."""
        is_overflow, direction = self.check_overflow(shape)
        if not is_overflow:
            return False

        st = shape.shape_type
        if st in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.TABLE,
                  MSO_SHAPE_TYPE.CHART):
            return self._adjust_by_scaling(shape, direction, max_scale_down)
        else:
            return self._adjust_by_moving(shape, direction)

    def _adjust_by_scaling(self, shape, direction, max_scale_down):
        """Scale down shape to fit within boundaries."""
        scale_x = 1.0
        scale_y = 1.0

        right = shape.left + shape.width
        bottom = shape.top + shape.height

        if right > self.slide_width:
            scale_x = self.slide_width / right

        if bottom > self.slide_height:
            scale_y = self.slide_height / bottom

        scale = min(scale_x, scale_y, max_scale_down)

        if scale < 1.0:
            shape.width = shape.width * scale
            shape.height = shape.height * scale
            return True
        return False

    def _adjust_by_moving(self, shape, direction):
        """Move shape to fit within boundaries."""
        moved = False

        right = shape.left + shape.width
        bottom = shape.top + shape.height

        if right > self.slide_width:
            shape.left = max(0, self.slide_width - shape.width)
            moved = True

        if bottom > self.slide_height:
            shape.top = max(0, self.slide_height - shape.height)
            moved = True

        if shape.left < 0:
            shape.left = 0
            moved = True

        if shape.top < 0:
            shape.top = 0
            moved = True

        return moved
