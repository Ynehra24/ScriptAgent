---
dataset_info:
  features:
  - name: image_url
    dtype: string
  - name: image
    dtype: image
  - name: instruction
    dtype: string
  - name: bbox
    sequence: float32
  - name: point
    sequence: float32
  - name: type
    dtype: string
  splits:
  - name: train
    num_bytes: 16591347652.088
    num_examples: 7496
  download_size: 327573839
  dataset_size: 16591347652.088
configs:
- config_name: default
  data_files:
  - split: train
    path: data/train-*
---

[Github](https://github.com/showlab/ShowUI/tree/main) | [arXiv](https://arxiv.org/abs/2411.17465) | [HF Paper](https://huggingface.co/papers/2411.17465) | [Spaces](https://huggingface.co/spaces/showlab/ShowUI) | [Datasets](https://huggingface.co/datasets/showlab/ShowUI-desktop-8K) | [Quick Start](https://huggingface.co/showlab/ShowUI-2B) 

**ShowUI-desktop-8K** is a UI-grounding dataset focused on PC-based grounding, with screenshots and annotations originally sourced from [OmniAct](https://huggingface.co/datasets/Writer/omniact). 

We utilize GPT-4o to augment the original annotations, enriching them with diverse attributes such as appearance, spatial relationships, and intended functionality.

You can use our [rewrite strategy code](https://github.com/showlab/ShowUI/blob/main/recaption.ipynb) to augment your own data.

![image/png](https://cdn-uploads.huggingface.co/production/uploads/64440be5af034cdfd69ca3a7/t6dZzpBdiDTHDxlke4Eva.png)

If you find our work helpful, please consider citing our paper.

```
@misc{lin2024showui,
      title={ShowUI: One Vision-Language-Action Model for GUI Visual Agent}, 
      author={Kevin Qinghong Lin and Linjie Li and Difei Gao and Zhengyuan Yang and Shiwei Wu and Zechen Bai and Weixian Lei and Lijuan Wang and Mike Zheng Shou},
      year={2024},
      eprint={2411.17465},
      archivePrefix={arXiv},
      primaryClass={cs.CV},
      url={https://arxiv.org/abs/2411.17465}, 
}
```