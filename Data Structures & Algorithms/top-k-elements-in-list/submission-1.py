class Solution:
    def topKFrequent(self, nums: List[int], k: int) -> List[int]:
        count={}
        for num in nums:
            count[num]=1+count.get(num,0) #the standard step to create a hash map, of which includes the element themselves and the frequency. Since we want to compare frequency, sorting is always a choice.
        arr=[]
        for num,cnt in count.items():
            arr.append([cnt,num])
        arr.sort()

        res=[]
        while len(res)<k:
            res.append(arr.pop()[1])
        return res

