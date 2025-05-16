from datetime import datetime

# Get current time
now = datetime.now()




def get_time_difference(start_time: str, end_time: str):
    # Parse HH:MM:SS into hours, minutes, seconds
    start_h, start_m, start_s = map(int, start_time.split(':'))
    end_h, end_m, end_s = map(int, end_time.split(':'))

    # Converting both times to total seconds
    start_total_seconds = start_h * 3600 + start_m * 60 + start_s
    end_total_seconds = end_h * 3600 + end_m * 60 + end_s

    # Calculating the difference in seconds
    difference_seconds = end_total_seconds - start_total_seconds

    # Handling negative difference
    if difference_seconds < 0:
        raise ValueError("End time must be after start time.")

    # Convert back to hours, minutes, seconds
    hours = difference_seconds // 3600
    minutes = (difference_seconds % 3600) // 60
    seconds = difference_seconds % 60

    return hours, minutes, seconds