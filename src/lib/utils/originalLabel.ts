const ORDINALS = ["첫 번째", "두 번째", "세 번째", "네 번째", "다섯 번째", "여섯 번째", "일곱 번째", "여덟 번째", "아홉 번째", "열 번째"];

export const getOriginalLabel = (index: number) => `${ORDINALS[index] ?? `${index + 1}번째`} 엑셀`;
