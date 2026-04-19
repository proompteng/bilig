export const WORKBOOK_RECT_SHADER = /* wgsl */ `
struct SurfaceUniforms {
  size: vec2f,
  origin: vec2f,
};

@group(0) @binding(0) var<uniform> surface: SurfaceUniforms;

struct VertexOut {
  @builtin(position) position: vec4f,
  @location(0) color: vec4f,
};

@vertex
fn vs_main(
  @location(0) quad: vec2f,
  @location(1) rect_origin: vec2f,
  @location(2) rect_size: vec2f,
  @location(3) rect_color: vec4f,
) -> VertexOut {
  let pixel = surface.origin + rect_origin + quad * rect_size;
  let clip = vec2f(
    (pixel.x / surface.size.x) * 2.0 - 1.0,
    1.0 - (pixel.y / surface.size.y) * 2.0,
  );

  var out: VertexOut;
  out.position = vec4f(clip, 0.0, 1.0);
  out.color = rect_color;
  return out;
}

@fragment
fn fs_main(@location(0) color: vec4f) -> @location(0) vec4f {
  return color;
}
`

export const WORKBOOK_TEXT_SHADER = /* wgsl */ `
struct SurfaceUniforms {
  size: vec2f,
  origin: vec2f,
};

@group(0) @binding(0) var<uniform> surface: SurfaceUniforms;
@group(0) @binding(1) var atlasSampler: sampler;
@group(0) @binding(2) var atlasTexture: texture_2d<f32>;

struct VertexOut {
  @builtin(position) position: vec4f,
  @location(0) uv: vec2f,
  @location(1) color: vec4f,
  @location(2) pixel: vec2f,
  @location(3) clip: vec4f,
};

@vertex
fn vs_main(
  @location(0) quad: vec2f,
  @location(1) rect_origin: vec2f,
  @location(2) rect_size: vec2f,
  @location(3) uv0: vec2f,
  @location(4) uv1: vec2f,
  @location(5) tint: vec4f,
  @location(6) clip_rect: vec4f,
) -> VertexOut {
  let local_pixel = rect_origin + quad * rect_size;
  let pixel = surface.origin + local_pixel;
  let clip = vec2f(
    (pixel.x / surface.size.x) * 2.0 - 1.0,
    1.0 - (pixel.y / surface.size.y) * 2.0,
  );
  let translated_clip = vec4f(
    clip_rect.x + surface.origin.x,
    clip_rect.y + surface.origin.y,
    clip_rect.z + surface.origin.x,
    clip_rect.w + surface.origin.y,
  );

  var out: VertexOut;
  out.position = vec4f(clip, 0.0, 1.0);
  out.uv = vec2f(
    mix(uv0.x, uv1.x, quad.x),
    mix(uv0.y, uv1.y, quad.y),
  );
  out.color = tint;
  out.pixel = pixel;
  out.clip = translated_clip;
  return out;
}

@fragment
fn fs_main(
  @location(0) uv: vec2f,
  @location(1) color: vec4f,
  @location(2) pixel: vec2f,
  @location(3) clip_rect: vec4f,
) -> @location(0) vec4f {
  if (pixel.x < clip_rect.x || pixel.y < clip_rect.y || pixel.x > clip_rect.z || pixel.y > clip_rect.w) {
    discard;
  }
  let sampled = textureSample(atlasTexture, atlasSampler, uv);
  return vec4f(color.rgb, color.a * sampled.a);
}
`
