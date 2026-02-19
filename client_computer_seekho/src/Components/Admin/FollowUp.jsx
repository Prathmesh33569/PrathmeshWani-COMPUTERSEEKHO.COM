function FollowUp() {
  const followUps = [
    {
      id: 1,
      enquirer: "XXXXXX",
      student: "AAAAAA",
      phone: "9812345678",
      course: "PG DAC",
      date: "22-Oct-22",
      staff: "Ruchita",
    },
    {
      id: 2,
      enquirer: "YYYYY",
      student: "BBBBBBBB",
      phone: "9812345679",
      course: "PG DBDA",
      date: "14-Oct-22",
      staff: "Ruchita",
    },
    {
      id: 3,
      enquirer: "ZZZZZZ",
      student: "CCCCCCC",
      phone: "9812345680",
      course: "MSCIT",
      date: "14-Oct-22",
      staff: "Ruchita",
    },
    {
      id: 4,
      enquirer: "PPPPPPPPP",
      student: "QQQQQQ",
      phone: "1234567890",
      course: "JAVA",
      date: "13-Oct-22",
      staff: "Snehal",
    },
  ];

  return (
    <div className="p-6">
      <h1 className="text-xl font-bold mb-4">Follow-Up</h1>

      <div className="overflow-x-auto">
        <table className="w-full border border-black text-sm">
          <thead>
            <tr className="bg-green-800 text-white">
              <th className="border border-black px-2 py-2">Eng.Id</th>
              <th className="border border-black px-2 py-2">Enquirer Name</th>
              <th className="border border-black px-2 py-2">Student Name</th>
              <th className="border border-black px-2 py-2">Phone</th>
              <th className="border border-black px-2 py-2">Course</th>
              <th className="border border-black px-2 py-2">Follow-up date</th>
              <th className="border border-black px-2 py-2">Staff Name</th>
              <th className="border border-black px-2 py-2">Action</th>
            </tr>
          </thead>

          <tbody>
            {followUps.map((item) => (
              <tr key={item.id} className="text-center">
                <td className="border border-black px-2 py-1">{item.id}</td>
                <td className="border border-black px-2 py-1">
                  {item.enquirer}
                </td>
                <td className="border border-black px-2 py-1">
                  {item.student}
                </td>
                <td className="border border-black px-2 py-1">{item.phone}</td>
                <td className="border border-black px-2 py-1">{item.course}</td>
                <td className="border border-black px-2 py-1">{item.date}</td>
                <td className="border border-black px-2 py-1">{item.staff}</td>
                <td className="border border-black px-2 py-1">
                  <button className="border border-black px-3 py-1 hover:bg-gray-200">
                    CALL
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Pagination */}
      <div className="flex gap-4 mt-4 font-semibold">
        <span className="cursor-pointer hover:underline">Begin</span>
        <span className="cursor-pointer hover:underline">Next</span>
        <span className="cursor-pointer hover:underline">Previous</span>
        <span className="cursor-pointer hover:underline">End</span>
      </div>
    </div>
  );
}

export default FollowUp;
