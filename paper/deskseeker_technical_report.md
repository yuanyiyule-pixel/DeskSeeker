# DeskSeeker: Toward Precise Desktop Coordinate Grounding with Multi-Stage Refinement and Majority Voting

**Ben James**  
NovaSerene  
amart@novaserene.com

## Abstract

Desktop UI grounding maps a natural-language target description to an actionable screen coordinate. In real desktop screenshots, this problem is difficult because targets are often small, densely packed, visually similar, and embedded in high-resolution interfaces. We present DeskSeeker, a training-free framework for desktop coordinate grounding based on multi-stage grid refinement and majority voting. DeskSeeker progressively narrows the search space from a coarse full-screen grid to neighborhood-level fine regions and then to a locally upscaled refinement crop. Across grounding stages, the framework launches parallel model calls and aggregates responses through voting in order to reduce instability from single-run predictions. The system returns one logical desktop coordinate for downstream automation without executing the action itself. We release a public implementation of DeskSeeker for Windows desktop screenshots and discuss why multi-stage refinement can, in principle, increase the upper bound of grounding precision compared with single-pass localization. The current version is a technical report and public system release; large-scale empirical evaluation is left to future work.

## 1. Introduction

Desktop UI grounding is the task of mapping a natural-language description of a user interface target to an actionable screen coordinate. This capability is important for desktop agents, GUI automation, assistive systems, remote operation pipelines, and human-in-the-loop workflows. Recent vision-language models have substantially improved general visual understanding, but reliable desktop grounding remains challenging in practice. Real desktop screenshots are often high-resolution, visually cluttered, and semantically dense. Targets such as taskbar icons, toolbar buttons, tabs, search boxes, and neighboring controls may be small, partially ambiguous, or visually similar to one another. As a result, a single-pass grounding attempt often suffers from coarse localization, instability, or confusion among nearby elements.

We present DeskSeeker, a training-free framework for desktop screenshot grounding that improves localization precision through multi-stage refinement and majority voting. Rather than relying on a single direct prediction, DeskSeeker progressively narrows the search space from a coarse full-screen grid to a neighborhood-level fine grid and then to a locally upscaled final refinement view. At each grounding stage, the framework launches multiple parallel model calls and aggregates responses through a voting procedure to reduce run-to-run variance. An optional review stage further verifies the final candidate before returning a logical desktop coordinate for downstream action. DeskSeeker is designed as a practical orchestration layer on top of general-purpose vision-language models, requiring no additional model training while remaining compatible with real desktop screenshots and interactive automation pipelines.

The contribution of this work is primarily methodological and systems-oriented. We do not claim a completed benchmark study in the current version. Instead, we formalize the DeskSeeker framework, provide the design rationale behind its staged refinement strategy, and release a public implementation. Our core claim is therefore not that DeskSeeker has already been comprehensively validated against all prior work, but that it provides a plausible and practically useful framework for pushing desktop screenshot grounding toward finer coordinate precision while remaining training-free and modular.

## 2. Related Work

### 2.1 GUI Grounding and GUI Agents

Recent work on GUI grounding and GUI agents has shown that large multimodal models can be adapted to screen understanding and action planning. Representative systems include SeeClick, CogAgent, and OS-Atlas, which study GUI grounding, action generation, and broader agent capabilities over desktop or mobile interfaces. These approaches demonstrate that screen interaction is a promising application area for multimodal models, but many of them rely on specialized model training, task-specific pretraining, or broader GUI datasets. DeskSeeker differs from this line by focusing on a training-free orchestration strategy that can be layered on top of general-purpose vision-language models.

### 2.2 Screen Annotation and Set-of-Mark Prompting

Another relevant line of work emphasizes visual annotation or set-of-mark prompting. In these approaches, the model is guided by explicit markers, tags, or labels rendered onto the screen image in order to improve reference resolution. This idea is closely related to DeskSeeker, which overlays grid labels on the image and uses those labels to support coarse-to-fine spatial localization. The difference is that DeskSeeker operationalizes this idea as a multi-stage desktop grounding pipeline with explicit refinement stages, local crops, and coordinate-space handling.

### 2.3 Coarse-to-Fine and Zoom-Based Grounding

Recent systems have also explored coarse-to-fine search, zoom-based refinement, or iterative grounding for screen tasks. This family of ideas is directly relevant to DeskSeeker. Our framework belongs to the same general direction: it starts with a global search over the full screenshot, then progressively narrows the search region and increases local resolution before making the final coordinate decision. DeskSeeker combines this coarse-to-fine design with parallel voting and an optional review stage in a single training-free orchestration pipeline.

## 3. Problem Setting

We consider the following problem. The input is a desktop screenshot and a natural-language description of one target element or region within that screenshot. The goal is to return a single actionable coordinate corresponding to the described target. In DeskSeeker, the returned coordinate is expressed in logical desktop coordinates rather than physical capture-pixel coordinates. The framework does not execute the click itself; it only produces the coordinate for downstream systems.

This problem setting reflects practical desktop-agent usage. A downstream automation system may issue the click, drag, or interaction after grounding is complete. Separating grounding from action execution keeps the system modular and allows the coordinate predictor to remain a general interface-layer component.

## 4. DeskSeeker Framework

### 4.1 Overview

DeskSeeker follows a three-stage grounding pipeline. First, the system captures or loads a desktop screenshot and overlays an adaptive coarse grid over the full screen. A vision-language model is then asked to identify the region that most likely contains the target. Second, the system crops the selected coarse region together with its neighboring context and applies a finer grid only to the selected area in order to refine the target location. Third, the system extracts the chosen fine cell and its surrounding local neighborhood, upsamples this local crop, overlays a final grid, and performs a last-stage refinement for precise coordinate selection.

Across grounding stages, DeskSeeker issues multiple parallel model queries and starts voting after a predefined number of successful responses. This design improves robustness against inconsistent single-run outputs while preserving a simple training-free workflow. The final output is one logical desktop coordinate, and the system never executes the click itself.

### 4.2 Stage 1: Coarse Grid Grounding

In the first stage, DeskSeeker overlays a full-screen coarse grid on the screenshot. The purpose of this stage is not to select a final coordinate, but to reduce the global search space to a manageable region. The coarse grid is adaptive within a bounded range so that the number of cells remains large enough for localization yet small enough for readable labels.

The model receives the labeled screenshot and is prompted to identify the grid cell most likely to contain the target described in natural language. Because desktop screenshots may contain multiple semantically similar targets, the description is intended to remain target-specific and avoid accidental over-specification of output formatting or hit-area policy. DeskSeeker also injects a click-success hint that biases the model toward stable actionable positions rather than decorative fragments.

### 4.3 Stage 2: Fine Grid Refinement

After selecting a coarse cell, the framework crops the chosen coarse cell together with neighboring coarse cells for context. It then applies a finer grid only within the selected coarse cell. The goal of this stage is to refine the search inside the previously selected region while preserving enough surrounding context to avoid semantic drift. If the model indicates that the target actually lies in a neighboring coarse region, DeskSeeker can shift the working region and repeat the second stage rather than restarting the full pipeline.

### 4.4 Stage 3: Local Upscaled Refinement

The third stage is designed for fine-grained coordinate localization. DeskSeeker extracts the stage-2 fine cell and its surrounding local neighborhood from the original crop, upsamples that local crop, and overlays another grid for final refinement. Upsampling before final grid annotation is important for small targets because it increases the visual separability of tiny elements and can reduce the chance that a final grid cell mixes multiple plausible click areas. The third stage returns the final local cell and the within-cell anchor used to derive the coordinate.

### 4.5 Parallel Voting and Optional Review

At each grounding stage, DeskSeeker launches multiple parallel model calls and begins voting after a fixed number of successful responses. The motivation is to reduce instability from single-run stochastic variation, transport-level failures, or inconsistent local interpretations. Instead of trusting a single response, the framework accepts the majority-supported candidate when the agreement threshold is met. An optional review stage can then verify the final candidate by drawing it back onto the full screenshot and asking the model whether the marked region matches the requested target.

### 4.6 Output Coordinate Definition

DeskSeeker returns one logical desktop coordinate. This coordinate is intended for downstream automation systems and is not itself an action. In practical deployment, the distinction between logical desktop coordinates and physical capture-pixel coordinates matters because some screenshot backends may initially produce physical-size captures that must be rescaled to match the logical desktop space used by actual GUI automation. DeskSeeker explicitly handles this distinction and returns the logical desktop coordinate as the final output contract.

## 5. Design Rationale

DeskSeeker is motivated by four design principles.

First, progressive search-space reduction can increase effective localization precision. A direct single-pass prediction over the entire screenshot forces the model to reason about semantic identity and exact spatial placement at once. By narrowing the candidate region stage by stage, the framework allows each stage to solve a smaller spatial problem.

Second, local upscaling before the final grid overlay can improve the theoretical upper bound of fine-grained localization. If a small icon or control occupies only a tiny set of pixels in the original crop, then drawing a final refinement grid over the unscaled crop can lead to mixed cells or visually ambiguous labels. Upscaling the local crop before the final annotation provides more visual separation and therefore a more favorable representation for precise grounding.

Third, majority voting can reduce variance. A single model response may fail due to transient ambiguity, local hallucination, or fragile reasoning about near-duplicate icons. Parallel calls with voting do not eliminate systematic bias, but they can reduce instability caused by run-to-run variation.

Fourth, an explicit verification stage can act as a posterior filter. Even if a candidate coordinate is plausible, a final review over the original screenshot can reject obviously wrong or unsafe selections before the output is returned.

Taken together, these design choices suggest why DeskSeeker may support finer coordinate grounding than a simpler single-pass workflow. The current report frames these arguments as design rationale rather than finalized empirical claims.

## 6. Public Implementation

We release a public implementation of DeskSeeker for Windows desktop screenshots. The release contains a runnable Node.js script, compact skill-style documentation, an MIT license, and a Zenodo-backed DOI for citation. The public version removes private key-file fallback logic and authenticates only through the `OPENROUTER_API_KEY` environment variable. The implementation returns a single logical desktop coordinate and is intended to serve as a practical reference baseline for future desktop grounding work.

## 7. Limitations and Current Scope

The current version of DeskSeeker has several limitations.

First, this report does not include a large-scale benchmark evaluation. We therefore do not present quantitative claims about success rate, coordinate error, latency-performance tradeoffs, or category-specific robustness. Large-scale empirical validation is left to future work.

Second, DeskSeeker currently targets Windows desktop screenshots and has not yet been systematically adapted to other operating systems or mobile UI environments.

Third, the framework depends on external vision-language model APIs. As a result, performance, latency, and cost may vary across models and providers.

Fourth, the method is still vulnerable to extreme ambiguity, missing visual evidence, hidden targets, and lookalike controls that remain semantically difficult even after multi-stage refinement.

Finally, the framework currently emphasizes practical coordinate grounding rather than full end-to-end GUI agent behavior. It does not execute actions and does not attempt to solve broader planning or interaction-loop problems.

## 8. Conclusion

We presented DeskSeeker, a training-free framework for desktop screenshot grounding based on multi-stage refinement and majority voting. DeskSeeker progressively reduces the search space from coarse full-screen localization to fine local refinement and returns one logical desktop coordinate without executing the action itself. The current contribution is a framework and public implementation rather than a completed benchmark study. We position DeskSeeker as a practical system design and a public starting point for future work on more precise desktop coordinate grounding with general-purpose vision-language models.

## Suggested Citation

If citing the software release directly, use the Zenodo DOI:

- DOI: `10.5281/zenodo.18906893`
- Record: `https://zenodo.org/records/18906893`
