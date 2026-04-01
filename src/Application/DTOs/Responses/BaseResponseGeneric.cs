namespace Application.DTOs.Responses;

public class BaseResponseGeneric<T> : BaseResponse
{
    public T? Data { get; set; }
}
