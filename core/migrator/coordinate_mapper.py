"""
CoordinateMapper: map object coordinates from source PPT to target PPT.

Handles different aspect ratios (4:3 → 16:9) by scaling and repositioning
objects to fit within the target slide dimensions.
"""

class CoordinateMapper:
    """Map coordinates from source to target slide dimensions."""

    def __init__(self, src_width, src_height, tgt_width, tgt_height):
        self.src_width = src_width
        self.src_height = src_height
        self.tgt_width = tgt_width
        self.tgt_height = tgt_height

        self._calculate_scale()

    def _calculate_scale(self):
        """Calculate scale factors for width and height."""
        self.scale_x = self.tgt_width / self.src_width
        self.scale_y = self.tgt_height / self.src_height
        self.min_scale = min(self.scale_x, self.scale_y)
        self.max_scale = max(self.scale_x, self.scale_y)

    def map_position(self, left, top, width, height, mode="fit_width") -> tuple:
        """
        Map position from source to target.
        mode:
          - "fit_width": scale to fit target width (height may overflow or have gaps)
          - "fit_height": scale to fit target height (width may overflow or have gaps)
          - "fit": scale to fit within both, centered
          - "stretch": stretch to fill entire target
        """
        if mode == "fit_width":
            scale = self.scale_x
        elif mode == "fit_height":
            scale = self.scale_y
        elif mode == "fit":
            scale = self.min_scale
        elif mode == "stretch":
            scale = (self.scale_x, self.scale_y)
        else:
            scale = self.scale_x

        if isinstance(scale, tuple):
            new_left = left * scale[0]
            new_top = top * scale[1]
            new_width = width * scale[0]
            new_height = height * scale[1]
        else:
            new_left = left * scale
            new_top = top * scale
            new_width = width * scale
            new_height = height * scale

        if mode == "fit" and not isinstance(scale, tuple):
            offset_x = (self.tgt_width - new_width) / 2
            offset_y = (self.tgt_height - new_height) / 2
            new_left += offset_x
            new_top += offset_y

        return new_left, new_top, new_width, new_height

    def map_shape(self, shape, mode="fit_width"):
        """Map all properties of a shape from source to target."""
        left, top, width, height = self.map_position(
            shape.left, shape.top, shape.width, shape.height, mode
        )
        return left, top, width, height
