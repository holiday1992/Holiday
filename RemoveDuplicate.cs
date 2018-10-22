namespace Array
{
    public class RemoveDuplicate
    {
        public int RemoveDuplicates(int[] nums)

        {
            int i = 0;
            nums[i++] = nums[0];
            if (i < nums.Length)

            {
                for (int j = 1; j < nums.Length; j++)
                {
                    if (nums[j] != nums[j - 1])
                    {
                        nums[i++] = nums[j];
                    }

                }

            }

            return i;

        }
    }
}