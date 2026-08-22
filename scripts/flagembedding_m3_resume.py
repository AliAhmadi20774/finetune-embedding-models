"""Compatibility entry point for resuming FlagEmbedding M3 training."""

from transformers import HfArgumentParser

from FlagEmbedding.finetune.embedder.encoder_only.m3 import (
    EncoderOnlyEmbedderM3DataArguments,
    EncoderOnlyEmbedderM3Model,
    EncoderOnlyEmbedderM3ModelArguments,
    EncoderOnlyEmbedderM3Runner,
    EncoderOnlyEmbedderM3TrainingArguments,
)


def main() -> None:
    # transformers 4.55 reads this standard PreTrainedModel attribute after
    # loading a checkpoint. FlagEmbedding's torch.nn.Module wrapper omits it.
    if not hasattr(EncoderOnlyEmbedderM3Model, "_keys_to_ignore_on_save"):
        EncoderOnlyEmbedderM3Model._keys_to_ignore_on_save = None

    parser = HfArgumentParser(
        (
            EncoderOnlyEmbedderM3ModelArguments,
            EncoderOnlyEmbedderM3DataArguments,
            EncoderOnlyEmbedderM3TrainingArguments,
        )
    )
    model_args, data_args, training_args = parser.parse_args_into_dataclasses()
    runner = EncoderOnlyEmbedderM3Runner(
        model_args=model_args,
        data_args=data_args,
        training_args=training_args,
    )
    runner.run()


if __name__ == "__main__":
    main()
