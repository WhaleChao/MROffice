/*!
 * MR Office - https://mroffice.ai
 * Copyright (C) 2026 MRNote. SPDX-License-Identifier: AGPL-3.0-only
 * Licensed under the GNU AGPL v3 with the additional terms in NOTICE.md.
 */
/**
 * glTF `KHR_materials_pbrSpecularGlossiness` for three's GLTFLoader.
 *
 * Office's GLTF exporter writes nearly every model PowerPoint inserts with
 * this extension, and requires it (model3d-truth: 19 of the corpus's 23
 * models). three's loader dropped it (r147), so without it such a model is
 * drawn from an empty `pbrMetallicRoughness`: white, fully metallic, fully
 * rough. This plugin draws the specular-glossiness model as the extension's
 * own specification states it (Khronos, archived extension):
 *
 *   diffuse colour  c_diff = diffuse.rgb * (1 - max(specular.r, .g, .b))
 *   F0              = specular
 *   roughness       = 1 - glossiness
 *
 * on three's physical lighting: a MeshStandardMaterial whose shader takes
 * F0 and roughness from the extension instead of metalness. The
 * specular-glossiness texture (sRGB specular in RGB, linear glossiness in
 * A) rides in the material's roughness map slot so three gives it its UV.
 *
 * Office marks a material whose factors are sRGB values
 * (`extras.MSFT_sRGBFactors`: true); they are made linear as read.
 */

import {
  Color,
  MeshStandardMaterial,
  SRGBColorSpace,
  type Material,
  type Texture,
} from "three";
import type { GLTFLoaderPlugin, GLTFParser } from "three/examples/jsm/loaders/GLTFLoader.js";

const EXTENSION = "KHR_materials_pbrSpecularGlossiness";

interface SpecGlossDefinition {
  diffuseFactor?: number[];
  diffuseTexture?: { index: number; texCoord?: number };
  specularFactor?: number[];
  glossinessFactor?: number;
  specularGlossinessTexture?: { index: number; texCoord?: number };
}

interface MaterialDefinition {
  extensions?: Record<string, unknown>;
  extras?: Record<string, unknown>;
}

function srgbToLinear(value: number): number {
  return value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4;
}

/** The extension's values for one material, linear. */
interface SpecGlossValues {
  specular: [number, number, number];
  glossiness: number;
  texture: Texture | null;
}

const SHADER_KEY = "mrai-spec-gloss-1";

/** Make `material` draw the specular-glossiness model with `values`. */
export function applySpecularGlossiness(material: MeshStandardMaterial, values: SpecGlossValues): void {
  material.metalness = 0;
  material.metalnessMap = null;
  material.roughness = 1;
  material.roughnessMap = values.texture;
  material.userData.specGloss = values;
  material.onBeforeCompile = (shader) => {
    shader.uniforms.sgSpecular = { value: new Color(values.specular[0], values.specular[1], values.specular[2]) };
    shader.uniforms.sgGlossiness = { value: values.glossiness };
    shader.fragmentShader = shader.fragmentShader
      .replace(
        "uniform float roughness;",
        "uniform float roughness;\nuniform vec3 sgSpecular;\nuniform float sgGlossiness;",
      )
      .replace(
        "#include <roughnessmap_fragment>",
        [
          "vec3 sgSpecularFactor = sgSpecular;",
          "float sgGlossinessFactor = sgGlossiness;",
          "#ifdef USE_ROUGHNESSMAP",
          "  vec4 sgTexel = texture2D( roughnessMap, vRoughnessMapUv );",
          "  sgSpecularFactor *= sgTexel.rgb;",
          "  sgGlossinessFactor *= sgTexel.a;",
          "#endif",
          "float roughnessFactor = 1.0 - sgGlossinessFactor;",
        ].join("\n"),
      )
      .replace(
        "#include <lights_physical_fragment>",
        [
          "#include <lights_physical_fragment>",
          "material.metalness = 0.0;",
          "material.diffuseContribution = diffuseColor.rgb * ( 1.0 - max( max( sgSpecularFactor.r, sgSpecularFactor.g ), sgSpecularFactor.b ) );",
          "material.specularColor = sgSpecularFactor;",
          "material.specularColorBlended = sgSpecularFactor;",
          "material.specularF90 = 1.0;",
        ].join("\n"),
      );
  };
  material.customProgramCacheKey = () => SHADER_KEY;
  material.needsUpdate = true;
}

/** The loader plugin: register with `loader.register((parser) => new SpecularGlossinessPlugin(parser))`. */
export class SpecularGlossinessPlugin implements GLTFLoaderPlugin {
  readonly name = EXTENSION;
  private readonly parser: GLTFParser;
  private readonly pending = new Map<number, Promise<SpecGlossValues>>();

  constructor(parser: GLTFParser) {
    this.parser = parser;
  }

  private definition(materialIndex: number): { material: MaterialDefinition; extension: SpecGlossDefinition } | null {
    const materials = (this.parser.json as { materials?: MaterialDefinition[] }).materials ?? [];
    const material = materials[materialIndex];
    const extension = material?.extensions?.[EXTENSION] as SpecGlossDefinition | undefined;
    return material && extension ? { material, extension } : null;
  }

  extendMaterialParams(materialIndex: number, params: Record<string, unknown>): Promise<unknown> | null {
    const found = this.definition(materialIndex);
    if (!found) return null;
    const { material, extension } = found;
    const srgbFactors = material.extras?.MSFT_sRGBFactors === true;
    const toLinear = (value: number) => (srgbFactors ? srgbToLinear(value) : value);
    const diffuse = extension.diffuseFactor ?? [1, 1, 1, 1];
    params.color = new Color(toLinear(diffuse[0] ?? 1), toLinear(diffuse[1] ?? 1), toLinear(diffuse[2] ?? 1));
    params.opacity = diffuse[3] ?? 1;
    params.metalness = 0;
    params.roughness = 1;
    const work: Promise<unknown>[] = [];
    if (extension.diffuseTexture) {
      work.push(this.parser.assignTexture(params, "map", extension.diffuseTexture, SRGBColorSpace));
    }
    const specular = extension.specularFactor ?? [1, 1, 1];
    const values: SpecGlossValues = {
      specular: [toLinear(specular[0] ?? 1), toLinear(specular[1] ?? 1), toLinear(specular[2] ?? 1)],
      glossiness: extension.glossinessFactor ?? 1,
      texture: null,
    };
    const textureDefinition = extension.specularGlossinessTexture;
    const ready = textureDefinition
      ? this.parser.getDependency("texture", textureDefinition.index).then((texture: Texture) => {
          const own = texture.clone();
          own.colorSpace = SRGBColorSpace;
          own.channel = textureDefinition.texCoord ?? 0;
          own.needsUpdate = true;
          values.texture = own;
          return values;
        })
      : Promise.resolve(values);
    this.pending.set(materialIndex, ready);
    work.push(ready);
    return Promise.all(work);
  }

  /** Called by the loader once the scene is built: give each material its shader. */
  async afterRoot(): Promise<void> {
    const associations = this.parser.associations as Map<unknown, { materials?: number }>;
    const done = new Set<Material>();
    for (const [object, reference] of associations) {
      if (!(object instanceof MeshStandardMaterial) || done.has(object)) continue;
      const index = reference?.materials;
      if (typeof index !== "number") continue;
      const pending = this.pending.get(index);
      if (!pending) continue;
      applySpecularGlossiness(object, await pending);
      done.add(object);
    }
  }
}
