import { atom, selector } from "recoil";

export const tagsListAtom = atom<string[]>({
  key: "tagList",
  default: [],
});

export const tagSelectedSelector = selector<string[]>({
  key: "tagSelected",
  get: ({ get }) => {
    const tagList = get(tagsListAtom);
    return tagList;
  },
});
