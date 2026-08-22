"""Evaluate dense retrieval against test positive and negative documents."""

from __future__ import annotations

import json

from evaluate_positive_only import create_parser, evaluate, validate_args


def main() -> None:
    parser = create_parser(
        "Evaluate BGE-M3 against all test positive and negative documents."
    )
    args = parser.parse_args()
    validate_args(parser, args)
    report = evaluate(
        args.test_file,
        args.model_name_or_path,
        args.output_dir,
        mode="positive_and_negative",
        devices=args.devices,
        use_fp16=args.fp16,
        encode_batch_size=args.encode_batch_size,
        search_batch_size=args.search_batch_size,
        query_max_length=args.query_max_length,
        passage_max_length=args.passage_max_length,
    )
    print(json.dumps(report, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
